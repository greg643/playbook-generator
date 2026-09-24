// Apple authenticates the coach; the GSS verifier supplies the same account key
// used by the native app. Existing playbooks keep their immutable local userId.
import { accountSessionVersion, createSessionCookie, generatePasswordResetToken,
  getSecret, getUser, hasRecentAuthentication, jsonNoStore, readBoundedUtf8Text,
  sha256Hex, validCredentialKey } from "./auth.js";
import { createJson, putJsonIfCurrent } from "./r2.js";

const TXN_COOKIE = "__Host-pb_apple_txn";
const PENDING_COOKIE = "__Host-pb_apple_pending";
const TTL = 600;
const KEY_RE = /^[a-f0-9]{16}$/;
const destinations = new Set(["/", "/editor", "/converter"]);
const bytes = value => new TextEncoder().encode(value);
const now = () => Math.floor(Date.now() / 1000);
export const appleKey = key => {
  if (!KEY_RE.test(key || "")) throw new Error("Invalid GSS identity");
  return `users/byapple/${key}.json`;
};

function encode(value) {
  return btoa(String.fromCharCode(...new Uint8Array(value))).replace(/\+/g, "-").replace(/\//g, "_").replace(/=+$/, "");
}
function decode(value) {
  return Uint8Array.from(atob(value.replace(/-/g, "+").replace(/_/g, "/")), c => c.charCodeAt(0));
}
function cookieValue(request, name) {
  return (request.headers.get("cookie") || "").split(";").map(p => p.trim())
    .find(p => p.startsWith(name + "="))?.slice(name.length + 1) || "";
}
function cookie(name, value, maxAge = TTL) {
  // Apple returns a cross-site POST, so the transaction cookie needs None.
  return `${name}=${value}; Path=/; HttpOnly; Secure; SameSite=${name === TXN_COOKIE ? "None" : "Lax"}; Max-Age=${maxAge}`;
}
async function aesKey(env) {
  return crypto.subtle.importKey("raw", await crypto.subtle.digest("SHA-256",
    bytes("playbook-apple-cookie-v1:" + await getSecret(env))), "AES-GCM", false, ["encrypt", "decrypt"]);
}
export async function sealApple(env, payload) {
  const iv = crypto.getRandomValues(new Uint8Array(12));
  const ciphertext = await crypto.subtle.encrypt({ name: "AES-GCM", iv }, await aesKey(env), bytes(JSON.stringify(payload)));
  const result = new Uint8Array(12 + ciphertext.byteLength);
  result.set(iv); result.set(new Uint8Array(ciphertext), 12);
  return encode(result);
}
async function readSealed(request, env, name, kind) {
  try {
    const raw = decode(cookieValue(request, name));
    const plain = await crypto.subtle.decrypt({ name: "AES-GCM", iv: raw.slice(0, 12) }, await aesKey(env), raw.slice(12));
    const value = JSON.parse(new TextDecoder().decode(plain));
    if (value.kind !== kind || value.origin !== new URL(request.url).origin ||
        !Number.isSafeInteger(value.iat) || value.iat > now() || now() - value.iat >= TTL) return null;
    return value;
  } catch { return null; }
}

export function appleConfigured(env, request) {
  if (env.APPLE_ENABLED !== "true" || !env.APPLE_SERVICES_ID || !env.APPLE_TEAM_ID ||
      !env.APPLE_KEY_ID || !env.APPLE_PRIVATE_KEY || !/^[a-f0-9]{64}$/i.test(env.SESSION_SECRET || "") ||
      !env.AUTH_STATE_BUCKET || !env.PLAYBOOK_BUCKET) return false;
  try {
    const origin = new URL(env.PUBLIC_ORIGIN);
    const api = new URL(env.GSS_API_BASE);
    return origin.protocol === "https:" && origin.origin === env.PUBLIC_ORIGIN &&
      new URL(request.url).origin === origin.origin && api.protocol === "https:" && !api.username && !api.password &&
      !api.search && !api.hash;
  } catch { return false; }
}

function redirect(location, cookies = []) {
  const headers = new Headers({ Location: location, "Cache-Control": "no-store", "Referrer-Policy": "no-referrer" });
  for (const value of cookies) headers.append("Set-Cookie", value);
  return new Response(null, { status: 303, headers });
}
function failure(reason = "failed") {
  return redirect("/?apple=" + reason, [cookie(TXN_COOKIE, "", 0), cookie(PENDING_COOKIE, "", 0)]);
}
class AppleExchangeError extends Error {
  constructor(code) { super("Identity exchange failed"); this.code = code; }
}
function callbackFailure(code) {
  // Codes are fixed literals from this module, never provider text or credentials.
  console.error(JSON.stringify({ event: "apple_callback_failed", code }));
  return failure("failed&apple_step=" + code);
}
function sameOrigin(request) {
  return request.headers.get("origin") === new URL(request.url).origin &&
    request.headers.get("content-type")?.split(";")[0].trim() === "application/json";
}

async function clientSecret(env) {
  const part = obj => encode(bytes(JSON.stringify(obj)));
  const input = part({ alg: "ES256", kid: env.APPLE_KEY_ID, typ: "JWT" }) + "." + part({
    iss: env.APPLE_TEAM_ID, sub: env.APPLE_SERVICES_ID, aud: "https://appleid.apple.com", iat: now(), exp: now() + 300,
  });
  const pem = env.APPLE_PRIVATE_KEY.replace(/-----[A-Z ]+-----/g, "").replace(/\s/g, "");
  const key = await crypto.subtle.importKey("pkcs8", decode(pem), { name: "ECDSA", namedCurve: "P-256" }, false, ["sign"]);
  return input + "." + encode(await crypto.subtle.sign({ name: "ECDSA", hash: "SHA-256" }, key, bytes(input)));
}

async function fetchJson(url, options, provider) {
  // Workerd rejects redirect: "error". Manual mode plus an explicit rejection
  // keeps identity credentials from ever being forwarded to a redirect target.
  const response = await fetch(url, { ...options, redirect: "manual", signal: AbortSignal.timeout(15000) });
  if (response.status >= 300 && response.status < 400) {
    await response.body?.cancel();
    throw new AppleExchangeError(provider + "_response");
  }
  const text = await readBoundedUtf8Text(response, 32768);
  if (text === null) throw new AppleExchangeError(provider + "_response");
  let body;
  try { body = JSON.parse(text); } catch { throw new AppleExchangeError(provider + "_response"); }
  if (!response.ok) {
    if (provider === "apple" && ["invalid_client", "unauthorized_client"].includes(body?.error)) {
      throw new AppleExchangeError("apple_client");
    }
    if (provider === "apple" && body?.error === "invalid_grant") throw new AppleExchangeError("apple_code");
    if (provider === "gss" && response.status === 401) throw new AppleExchangeError("gss_rejected");
    if (response.status === 429) throw new AppleExchangeError(provider + "_rate_limit");
    throw new AppleExchangeError(provider + "_response");
  }
  return body;
}

async function consume(env, purpose, nonce) {
  if (!/^[A-Za-z0-9_-]{43}$/.test(nonce || "")) return false;
  // Dedicated bucket with a one-day lifecycle. Keep replay markers out of
  // account/job storage; the sealed cookies expire after ten minutes.
  return await createJson({ PLAYBOOK_BUCKET: env.AUTH_STATE_BUCKET },
    `apple-once/${await sha256Hex(purpose + ":" + nonce)}.json`, { usedAt: now() }) !== null;
}

export async function findAppleAccount(env, key) {
  const indexKey = appleKey(key);
  const object = await env.PLAYBOOK_BUCKET.get(indexKey);
  if (!object) return null;
  const index = await object.json();
  if (index.deletedAt) return null;
  if (index.kind === "link") {
    if (!validCredentialKey(index.credentialKey) || !index.credentialKey.startsWith("users/byemail/")) throw new Error("Invalid identity mapping");
    const target = await env.PLAYBOOK_BUCKET.get(index.credentialKey);
    const record = target && await target.json();
    if (!record || record.userId !== index.userId || record.appleAccountKey !== key) return null;
    return { record, credentialKey: index.credentialKey };
  }
  if (index.appleAccountKey !== key || typeof index.userId !== "string") throw new Error("Invalid Apple account");
  return { record: index, credentialKey: indexKey };
}

async function sessionFor(env, account, allowDisabled = false) {
  if ((!allowDisabled && account.record.disabledAt) || account.record.deletedAt) throw new Error("Account unavailable");
  return createSessionCookie(account.record.userId, account.record.email || "", env,
    accountSessionVersion(account.record), { credentialKey: account.credentialKey, appleAccountKey: account.record.appleAccountKey });
}

export async function startApple({ request, env }) {
  if (!appleConfigured(env, request)) return failure("unavailable");
  const url = new URL(request.url);
  const returnTo = destinations.has(url.searchParams.get("returnTo")) ? url.searchParams.get("returnTo") : "/editor";
  const state = generatePasswordResetToken();
  const rawNonce = generatePasswordResetToken();
  const txn = await sealApple(env, { kind: "transaction", origin: url.origin, iat: now(), state, rawNonce, returnTo });
  const authorize = new URL("https://appleid.apple.com/auth/authorize");
  authorize.search = new URLSearchParams({ client_id: env.APPLE_SERVICES_ID,
    redirect_uri: env.PUBLIC_ORIGIN + "/api/auth/apple/callback", response_type: "code",
    response_mode: "form_post", scope: "name", state, nonce: await sha256Hex(rawNonce) }).toString();
  return redirect(authorize.toString(), [cookie(TXN_COOKIE, txn), cookie(PENDING_COOKIE, "", 0)]);
}

export async function callbackApple({ request, env }) {
  let stage = "callback_request";
  try {
    if (!appleConfigured(env, request)) return failure("unavailable");
    if (request.headers.get("content-type")?.split(";")[0] !== "application/x-www-form-urlencoded") return callbackFailure(stage);
    const text = await readBoundedUtf8Text(request, 16384);
    if (text === null) return callbackFailure(stage);
    const form = new URLSearchParams(text);
    if (!cookieValue(request, TXN_COOKIE)) return callbackFailure("transaction_missing");
    const txn = await readSealed(request, env, TXN_COOKIE, "transaction");
    if (!txn) return callbackFailure("transaction_invalid");
    if (form.get("state") !== txn.state || !destinations.has(txn.returnTo)) return callbackFailure("state_mismatch");
    stage = "replay_storage";
    if (!await consume(env, "transaction", txn.state)) return callbackFailure("transaction_used");
    if (form.has("error")) return failure("cancelled");
    const code = form.get("code");
    if (!code || code.length > 4096) return callbackFailure("code_missing");
    stage = "signing_key";
    const secret = await clientSecret(env);
    stage = "apple_exchange";
    const apple = await fetchJson("https://appleid.apple.com/auth/token", {
      method: "POST", headers: { "Content-Type": "application/x-www-form-urlencoded" },
      body: new URLSearchParams({ client_id: env.APPLE_SERVICES_ID, client_secret: secret,
        code, grant_type: "authorization_code", redirect_uri: env.PUBLIC_ORIGIN + "/api/auth/apple/callback" }),
    }, "apple");
    if (typeof apple?.id_token !== "string" || apple.id_token.length > 16384) return callbackFailure("apple_response");
    let displayName;
    try {
      const name = JSON.parse(form.get("user") || "{}").name;
      displayName = [name?.firstName, name?.lastName].filter(v => typeof v === "string").join(" ").trim().slice(0, 80) || undefined;
    } catch { /* Apple sends the name only on first consent. */ }
    // GSS validates Apple's signature, issuer, audience, expiry and hashed nonce.
    // Never derive identity from an unverified client JWT or email address.
    stage = "gss_exchange";
    const verified = await fetchJson(env.GSS_API_BASE.replace(/\/$/, "") + "/v2/auth/apple", {
      method: "POST", headers: { "Content-Type": "application/json" },
      body: JSON.stringify({ identity_token: apple.id_token, raw_nonce: txn.rawNonce, display_name: displayName }),
    }, "gss");
    if (!KEY_RE.test(verified?.account_key || "") || !verified.session_token) return callbackFailure("gss_response");
    displayName = typeof verified.display_name === "string" && verified.display_name.trim()
      ? verified.display_name.trim().slice(0, 80) : (displayName || "Coach");
    stage = "account_lookup";
    const existing = await findAppleAccount(env, verified.account_key);
    stage = "session_create";
    if (existing) return redirect(existing.record.disabledAt ? "/apple-delete" : txn.returnTo,
      [await sessionFor(env, existing, true), cookie(TXN_COOKIE, "", 0), cookie(PENDING_COOKIE, "", 0)]);
    const pending = await sealApple(env, { kind: "pending", origin: txn.origin, iat: now(),
      nonce: generatePasswordResetToken(), appleAccountKey: verified.account_key, displayName, returnTo: txn.returnTo });
    return redirect("/apple-welcome", [cookie(TXN_COOKIE, "", 0), cookie(PENDING_COOKIE, pending)]);
  } catch (error) {
    // Do not log tokens, authorization codes, cookies or provider response bodies.
    return callbackFailure(error instanceof AppleExchangeError ? error.code : stage);
  }
}

export async function pendingApple({ request, env }) {
  if (!appleConfigured(env, request)) return jsonNoStore({ error: "Apple sign-in is unavailable" }, { status: 503 });
  const pending = await readSealed(request, env, PENDING_COOKIE, "pending");
  if (!pending) return jsonNoStore({ error: "Please continue with Apple again." }, { status: 401 });
  const user = await getUser(request, env);
  return jsonNoStore({ displayName: pending.displayName,
    account: user ? { userId: user.userId, email: user.email, displayName: user.displayName, recent: hasRecentAuthentication(user) } : null });
}

export async function finishApple({ request, env }) {
  if (!appleConfigured(env, request)) return jsonNoStore({ error: "Apple sign-in is unavailable" }, { status: 503 });
  if (!sameOrigin(request)) return jsonNoStore({ error: "Invalid request origin" }, { status: 403 });
  try {
    const text = await readBoundedUtf8Text(request, 2048);
    if (text === null) return jsonNoStore({ error: "Request too large" }, { status: 413 });
    let body;
    try { body = JSON.parse(text); } catch { return jsonNoStore({ error: "Invalid JSON" }, { status: 400 }); }
    if (!body || typeof body !== "object" || Array.isArray(body)) return jsonNoStore({ error: "Invalid request" }, { status: 400 });
    const pending = await readSealed(request, env, PENDING_COOKIE, "pending");
    if (!pending || !KEY_RE.test(pending.appleAccountKey) || !destinations.has(pending.returnTo)) {
      return jsonNoStore({ error: "Please continue with Apple again." }, { status: 401 });
    }
    if (!["new", "link"].includes(body.mode)) return jsonNoStore({ error: "Choose an account option" }, { status: 400 });
    const user = await getUser(request, env);
    if (body.mode === "link" && (!user || !hasRecentAuthentication(user) || user.userId !== body.userId)) {
      return jsonNoStore({ error: "Sign into your existing playbook account again, then link Apple." }, { status: 401 });
    }
    if (body.mode === "new" && user) return jsonNoStore({ error: "You are signed in. Link your existing playbooks instead." }, { status: 409 });
    if (!await consume(env, "finish", pending.nonce)) return jsonNoStore({ error: "This Apple sign-in was already used. Please continue with Apple again." }, { status: 409 });
    const key = appleKey(pending.appleAccountKey);
    const indexObject = await env.PLAYBOOK_BUCKET.get(key);
    let index = indexObject && await indexObject.json();
    if (index?.kind === "link" && validCredentialKey(index.credentialKey)) {
      const targetObject = await env.PLAYBOOK_BUCKET.get(index.credentialKey);
      const target = targetObject && await targetObject.json();
      // A failed CAS can leave a reservation behind. Reclaim it only when the
      // original immutable account is gone or committed to a different Apple ID.
      if (!target || target.deletedAt || target.userId !== index.userId ||
          (target.appleAccountKey && target.appleAccountKey !== pending.appleAccountKey)) index = null;
    }
    let account;
    if (body.mode === "new") {
      if (index && !index.deletedAt) return jsonNoStore({ error: "Apple is already connected to a playbook account. Continue with Apple to open it." }, { status: 409 });
      const record = { userId: crypto.randomUUID(), email: "", displayName: pending.displayName,
        appleAccountKey: pending.appleAccountKey, sessionVersion: 1, createdAt: new Date().toISOString() };
      if (await putJsonIfCurrent(env, key, record, indexObject) === null) return jsonNoStore({ error: "Account changed. Continue with Apple again." }, { status: 409 });
      account = { record, credentialKey: key };
    } else {
      if ((user.account.appleAccountKey && user.account.appleAccountKey !== pending.appleAccountKey) ||
          (index && !index.deletedAt && (index.userId !== user.userId || index.credentialKey !== user.credentialKey))) {
        return jsonNoStore({ error: "These accounts are already linked elsewhere. No playbooks were moved." }, { status: 409 });
      }
      if (!index || index.deletedAt) {
        const linked = await putJsonIfCurrent(env, key, { kind: "link", userId: user.userId, credentialKey: user.credentialKey }, indexObject);
        if (linked === null) return jsonNoStore({ error: "Apple was linked in another request. Please try again." }, { status: 409 });
      }
      const record = { ...user.account, appleAccountKey: pending.appleAccountKey,
        displayName: pending.displayName, sessionVersion: accountSessionVersion(user.account) + 1 };
      // Apple is an additional sign-in method. Keep the existing password and
      // recovery code, but invalidate outstanding reset links and old sessions
      // at this explicit security-sensitive account change.
      delete record.passwordReset;
      if (await putJsonIfCurrent(env, user.credentialKey, record, user.accountObject) === null) {
        return jsonNoStore({ error: "Your account changed while linking. Sign in and link Apple again; your plays are safe." }, { status: 409 });
      }
      account = { record, credentialKey: user.credentialKey };
    }
    const headers = new Headers();
    headers.append("Set-Cookie", await sessionFor(env, account));
    headers.append("Set-Cookie", cookie(PENDING_COOKIE, "", 0));
    return jsonNoStore({ ok: true, returnTo: pending.returnTo }, { headers });
  } catch {
    console.error("Apple account setup could not complete");
    return jsonNoStore({ error: "Could not finish account setup. Please continue with Apple again." }, { status: 503 });
  }
}
