// Shared auth helpers for Pages Functions. Web Crypto only, no libraries.

const SESSION_COOKIE = "pb_session";
const SESSION_TTL_SECONDS = 30 * 24 * 60 * 60; // 30 days
const RECENT_AUTH_SECONDS = 10 * 60;
const EMAIL_RE = /^[^\s@]+@[^\s@]+\.[^\s@]+$/;

export const PASSWORD_ITERATIONS = 600000;

export function jsonNoStore(data, init = {}) {
  const headers = new Headers(init.headers || {});
  headers.set("Cache-Control", "private, no-store");
  return Response.json(data, { ...init, headers });
}

// Reject a declared oversized body before Workers buffers it through text() or
// formData(). Callers still enforce the actual decoded byte count afterward,
// because Content-Length can be absent.
export function requestBodyTooLarge(request, maxBytes) {
  const raw = request.headers.get("content-length");
  if (!raw) return false;
  return !/^\d+$/.test(raw) || Number(raw) > maxBytes;
}

export function utf8ByteLength(text) {
  return new TextEncoder().encode(text).length;
}

export function normalizeEmail(email) {
  return String(email || "").trim().toLowerCase();
}

export function isValidEmail(email) {
  return typeof email === "string" && email.length <= 254 && EMAIL_RE.test(email);
}

function bytesToHex(bytes) {
  let hex = "";
  for (const b of bytes) hex += b.toString(16).padStart(2, "0");
  return hex;
}

function hexToBytes(hex) {
  const bytes = new Uint8Array(hex.length / 2);
  for (let i = 0; i < bytes.length; i++) {
    bytes[i] = parseInt(hex.slice(i * 2, i * 2 + 2), 16);
  }
  return bytes;
}

function b64urlEncode(bytes) {
  let bin = "";
  for (const b of bytes) bin += String.fromCharCode(b);
  return btoa(bin).replace(/\+/g, "-").replace(/\//g, "_").replace(/=+$/, "");
}

function b64urlDecode(str) {
  let s = str.replace(/-/g, "+").replace(/_/g, "/");
  while (s.length % 4) s += "=";
  const bin = atob(s);
  const bytes = new Uint8Array(bin.length);
  for (let i = 0; i < bin.length; i++) bytes[i] = bin.charCodeAt(i);
  return bytes;
}

function timingSafeEqual(a, b) {
  if (a.length !== b.length) return false;
  let diff = 0;
  for (let i = 0; i < a.length; i++) diff |= a[i] ^ b[i];
  return diff === 0;
}

export async function sha256Hex(text) {
  const digest = await crypto.subtle.digest("SHA-256", new TextEncoder().encode(text));
  return bytesToHex(new Uint8Array(digest));
}

export function emailKey(email) {
  return sha256Hex(email).then((hex) => `users/byemail/${hex}.json`);
}

export async function getSecret(env) {
  if (env.SESSION_SECRET !== undefined) {
    const configured = String(env.SESSION_SECRET).trim();
    if (!/^[0-9a-f]{64}$/i.test(configured)) {
      throw new Error("SESSION_SECRET must be exactly 64 hexadecimal characters");
    }
    return configured;
  }

  const obj = await env.PLAYBOOK_BUCKET.get("auth/secret");
  if (obj) {
    const stored = (await obj.text()).trim();
    if (!/^[0-9a-f]{64}$/i.test(stored)) throw new Error("Invalid stored session secret");
    return stored;
  }
  const bytes = new Uint8Array(32);
  crypto.getRandomValues(bytes);
  const secret = bytesToHex(bytes);
  const created = await env.PLAYBOOK_BUCKET.put("auth/secret", secret, {
    onlyIf: { etagDoesNotMatch: "*" },
    httpMetadata: { contentType: "text/plain" },
  });
  if (created !== null) return secret;

  // Another request initialized the secret first. Use that winner rather than
  // returning a key that was never stored.
  const winner = await env.PLAYBOOK_BUCKET.get("auth/secret");
  if (!winner) throw new Error("Session secret initialization failed");
  const stored = (await winner.text()).trim();
  if (!/^[0-9a-f]{64}$/i.test(stored)) throw new Error("Invalid stored session secret");
  return stored;
}

export async function hashPassword(password, saltHex, iterations) {
  const key = await crypto.subtle.importKey(
    "raw",
    new TextEncoder().encode(password),
    "PBKDF2",
    false,
    ["deriveBits"]
  );
  const bits = await crypto.subtle.deriveBits(
    { name: "PBKDF2", hash: "SHA-256", salt: hexToBytes(saltHex), iterations },
    key,
    256
  );
  return bytesToHex(new Uint8Array(bits));
}

export function generateSaltHex() {
  const bytes = new Uint8Array(16);
  crypto.getRandomValues(bytes);
  return bytesToHex(bytes);
}

const RECOVERY_ITERATIONS = 100000;

// Recovery codes are 20 hex chars shown as XXXX-XXXX-XXXX-XXXX-XXXX.
// Verification is case- and dash-insensitive: normalize before hashing.
export function generateRecoveryCode() {
  const bytes = new Uint8Array(10);
  crypto.getRandomValues(bytes);
  const hex = bytesToHex(bytes).toUpperCase();
  const groups = [];
  for (let i = 0; i < hex.length; i += 4) groups.push(hex.slice(i, i + 4));
  return groups.join("-");
}

export function normalizeRecoveryCode(code) {
  return String(code || "").toLowerCase().replace(/[^0-9a-f]/g, "");
}

// Fresh recovery code plus the PBKDF2 fields to persist on the user record.
export async function createRecoveryFields() {
  const recoveryCode = generateRecoveryCode();
  const recoverySalt = generateSaltHex();
  const recoveryHash = await hashPassword(
    normalizeRecoveryCode(recoveryCode),
    recoverySalt,
    RECOVERY_ITERATIONS
  );
  return {
    recoveryCode,
    fields: { recoverySalt, recoveryIterations: RECOVERY_ITERATIONS, recoveryHash },
  };
}

// Constant-time compare of two hex digest strings.
export function constantTimeEqualHex(a, b) {
  if (typeof a !== "string" || typeof b !== "string" || a.length === 0 || b.length === 0) {
    return false;
  }
  let diff = a.length === b.length ? 0 : 1;
  for (let i = 0; i < a.length; i++) {
    diff |= a.charCodeAt(i) ^ b.charCodeAt(i % b.length);
  }
  return diff === 0;
}

async function hmacSign(secret, payload) {
  const key = await crypto.subtle.importKey(
    "raw",
    hexToBytes(secret),
    { name: "HMAC", hash: "SHA-256" },
    false,
    ["sign"]
  );
  const sig = await crypto.subtle.sign("HMAC", key, new TextEncoder().encode(payload));
  return new Uint8Array(sig);
}

function accountSessionVersion(record) {
  return Number.isSafeInteger(record && record.sessionVersion) && record.sessionVersion >= 1
    ? record.sessionVersion
    : 1;
}

export async function createSessionCookie(userId, email, env, sessionVersion = 1) {
  const secret = await getSecret(env);
  const issuedAt = Math.floor(Date.now() / 1000);
  const exp = issuedAt + SESSION_TTL_SECONDS;
  const payload = b64urlEncode(
    new TextEncoder().encode(
      JSON.stringify({ uid: userId, em: normalizeEmail(email), sv: sessionVersion, iat: issuedAt, exp })
    )
  );
  const sig = b64urlEncode(await hmacSign(secret, payload));
  const token = `${payload}.${sig}`;
  return `${SESSION_COOKIE}=${token}; HttpOnly; Secure; SameSite=Lax; Path=/; Max-Age=${SESSION_TTL_SECONDS}`;
}

export function clearSessionCookie() {
  return `${SESSION_COOKIE}=; HttpOnly; Secure; SameSite=Lax; Path=/; Max-Age=0`;
}

export async function getUser(request, env, { allowDisabled = false } = {}) {
  const cookieHeader = request.headers.get("cookie") || "";
  let token = null;
  for (const part of cookieHeader.split(";")) {
    const eq = part.indexOf("=");
    if (eq === -1) continue;
    if (part.slice(0, eq).trim() === SESSION_COOKIE) {
      token = part.slice(eq + 1).trim();
      break;
    }
  }
  if (!token) return null;

  const dot = token.indexOf(".");
  if (dot === -1) return null;
  const payload = token.slice(0, dot);
  const sigPart = token.slice(dot + 1);

  try {
    const secret = await getSecret(env);
    const expected = await hmacSign(secret, payload);
    const given = b64urlDecode(sigPart);
    if (!timingSafeEqual(expected, given)) return null;

    const data = JSON.parse(new TextDecoder().decode(b64urlDecode(payload)));
    if (typeof data.uid !== "string" || typeof data.exp !== "number") return null;
    if (data.exp <= Math.floor(Date.now() / 1000)) return null;
    const email = normalizeEmail(data.em);
    if (!isValidEmail(email)) return null;

    // A valid HMAC is not enough: bind the cookie to the current immutable
    // account ID and revocation version. Missing versions on legacy accounts
    // and cookies are treated as version 1 so existing users stay signed in.
    const accountObject = await env.PLAYBOOK_BUCKET.get(await emailKey(email));
    if (!accountObject) return null;
    const account = await accountObject.json();
    const cookieVersion = Number.isSafeInteger(data.sv) && data.sv >= 1 ? data.sv : 1;
    if (
      !account ||
      account.userId !== data.uid ||
      normalizeEmail(account.email) !== email ||
      accountSessionVersion(account) !== cookieVersion ||
      (account.disabledAt && !allowDisabled)
    ) {
      return null;
    }

    return {
      userId: account.userId,
      email,
      sessionVersion: cookieVersion,
      authenticatedAt: Number.isSafeInteger(data.iat) ? data.iat : 0,
      account,
      accountObject,
    };
  } catch (err) {
    return null;
  }
}

// Re-check account liveness around mutations that write outside the account
// record. This deliberately ignores sessionVersion: password recovery may
// revoke a session while one of its already-authorized saves is finishing,
// but only deletion/identity replacement should make that write self-clean.
export async function isAccountActive(env, user) {
  if (!user || typeof user.email !== "string" || typeof user.userId !== "string") {
    return false;
  }
  const object = await env.PLAYBOOK_BUCKET.get(await emailKey(user.email));
  if (!object) return false;
  try {
    const account = await object.json();
    return (
      account &&
      account.userId === user.userId &&
      normalizeEmail(account.email) === user.email &&
      !account.disabledAt
    );
  } catch (error) {
    console.error("Could not verify account liveness:", error);
    return false;
  }
}

export function hasRecentAuthentication(user, nowSeconds = Math.floor(Date.now() / 1000)) {
  return (
    user &&
    Number.isSafeInteger(user.authenticatedAt) &&
    user.authenticatedAt > 0 &&
    user.authenticatedAt <= nowSeconds + 60 &&
    nowSeconds - user.authenticatedAt <= RECENT_AUTH_SECONDS
  );
}

// Returns { user, response }: on success user is set; on failure user is null
// and response is a ready-to-return 401.
export async function requireUser(context, options) {
  const user = await getUser(context.request, context.env, options);
  if (!user) {
    return { user: null, response: jsonNoStore({ error: "Not signed in" }, { status: 401 }) };
  }
  return { user, response: null };
}
