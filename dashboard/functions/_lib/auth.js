// Shared auth helpers for Pages Functions. Web Crypto only, no libraries.

const SESSION_COOKIE = "pb_session";
const SESSION_TTL_SECONDS = 30 * 24 * 60 * 60; // 30 days
const EMAIL_RE = /^[^\s@]+@[^\s@]+\.[^\s@]+$/;

export function jsonNoStore(data, init = {}) {
  const headers = new Headers(init.headers || {});
  headers.set("Cache-Control", "no-store");
  return Response.json(data, { ...init, headers });
}

export function normalizeEmail(email) {
  return String(email || "").trim().toLowerCase();
}

export function isValidEmail(email) {
  return EMAIL_RE.test(email);
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
  const obj = await env.PLAYBOOK_BUCKET.get("auth/secret");
  if (obj) {
    return (await obj.text()).trim();
  }
  const bytes = new Uint8Array(32);
  crypto.getRandomValues(bytes);
  const secret = bytesToHex(bytes);
  await env.PLAYBOOK_BUCKET.put("auth/secret", secret, {
    httpMetadata: { contentType: "text/plain" },
  });
  return secret;
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

export async function createSessionCookie(userId, email, env) {
  const secret = await getSecret(env);
  const exp = Math.floor(Date.now() / 1000) + SESSION_TTL_SECONDS;
  const payload = b64urlEncode(
    new TextEncoder().encode(JSON.stringify({ uid: userId, em: email, exp }))
  );
  const sig = b64urlEncode(await hmacSign(secret, payload));
  const token = `${payload}.${sig}`;
  return `${SESSION_COOKIE}=${token}; HttpOnly; Secure; SameSite=Lax; Path=/; Max-Age=${SESSION_TTL_SECONDS}`;
}

export function clearSessionCookie() {
  return `${SESSION_COOKIE}=; HttpOnly; Secure; SameSite=Lax; Path=/; Max-Age=0`;
}

export async function getUser(request, env) {
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
    return { userId: data.uid, email: typeof data.em === "string" ? data.em : "" };
  } catch (err) {
    return null;
  }
}

// Returns { user, response }: on success user is set; on failure user is null
// and response is a ready-to-return 401.
export async function requireUser(context) {
  const user = await getUser(context.request, context.env);
  if (!user) {
    return { user: null, response: jsonNoStore({ error: "Not signed in" }, { status: 401 }) };
  }
  return { user, response: null };
}
