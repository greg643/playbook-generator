// The editor's existing HttpOnly session is the handle, not a GSS credential.
// Encrypted grants live only in the private, one-day-lifecycle auth bucket.
import { getSecret, getUser, readBoundedUtf8Text, sha256Hex } from './auth.js';
import { AppAccessError } from './app-access.js';
import { validGSSBearer } from './gss-bearer.js';
export { validGSSBearer } from './gss-bearer.js';

const SCHEMA = 'gss.playbook-web-session.v1';
const bytes = value => new TextEncoder().encode(value);
const hex = value => [...new Uint8Array(value)].map(n => n.toString(16).padStart(2, '0')).join('');
const unhex = value => Uint8Array.from(value.match(/../g), n => parseInt(n, 16));
const now = () => Math.floor(Date.now() / 1000);
const expired = () => new AppAccessError(401, 'Continue with Apple again to publish to your Team');

function sessionHandle(cookie) {
  const parts = (cookie || '').split(';').map(p => p.trim()).filter(p => p.startsWith('pb_session='));
  if (parts.length !== 1 || !/^pb_session=[A-Za-z0-9_-]+\.[A-Za-z0-9_-]+$/.test(parts[0]) || parts[0].length > 4096) throw expired();
  return parts[0];
}
async function sessionKey(cookie) { return `app-web-session/${await sha256Hex(sessionHandle(cookie))}.json`; }
async function pendingKey(nonce) {
  if (!/^[A-Za-z0-9_-]{43}$/.test(nonce || '')) throw expired();
  return `app-web-pending/${await sha256Hex(nonce)}.json`;
}
async function aesKey(env) {
  return crypto.subtle.importKey('raw', await crypto.subtle.digest('SHA-256',
    bytes(SCHEMA + ':' + await getSecret(env))), 'AES-GCM', false, ['encrypt', 'decrypt']);
}
async function write(env, key, value) {
  if (!env.AUTH_STATE_BUCKET) throw new AppAccessError(503, 'Team publishing is temporarily unavailable');
  const iv = crypto.getRandomValues(new Uint8Array(12));
  const sealed = await crypto.subtle.encrypt({ name: 'AES-GCM', iv, additionalData: bytes(key) },
    await aesKey(env), bytes(JSON.stringify({ schema: SCHEMA, ...value })));
  await env.AUTH_STATE_BUCKET.put(key, JSON.stringify({ iv: hex(iv), sealed: hex(sealed) }),
    { httpMetadata: { contentType: 'application/json', cacheControl: 'private, no-store' } });
}
async function read(env, key, origin, ttl) {
  if (!env.AUTH_STATE_BUCKET) throw new AppAccessError(503, 'Team publishing is temporarily unavailable');
  const object = await env.AUTH_STATE_BUCKET.get(key);
  if (!object) throw expired();
  try {
    const text = await readBoundedUtf8Text(new Response(object.body), 65536);
    const value = text && JSON.parse(text);
    if (!/^[a-f0-9]{24}$/.test(value?.iv || '') || !/^(?:[a-f0-9]{2}){16,24000}$/.test(value?.sealed || '')) throw expired();
    const plain = await crypto.subtle.decrypt({ name: 'AES-GCM', iv: unhex(value.iv), additionalData: bytes(key) },
      await aesKey(env), unhex(value.sealed));
    const grant = JSON.parse(new TextDecoder().decode(plain));
    if (grant.schema !== SCHEMA || grant.origin !== origin || !Number.isSafeInteger(grant.issuedAt) ||
        grant.issuedAt > now() || now() - grant.issuedAt >= ttl || !validGSSBearer(grant.bearer)) throw expired();
    return grant;
  } catch { throw expired(); }
}

// New-account consent can take a second request. Its encrypted cookie carries
// only the nonce; even a maximum-size GSS token never expands browser cookies.
export async function savePendingAppSession(env, origin, nonce, accountKey, bearer) {
  if (!validGSSBearer(bearer) || !/^[a-f0-9]{16}$/.test(accountKey)) throw expired();
  await write(env, await pendingKey(nonce), { origin, accountKey, bearer, issuedAt: now() });
}
export async function readPendingAppSession(env, origin, nonce, accountKey) {
  const grant = await read(env, await pendingKey(nonce), origin, 600);
  if (grant.accountKey !== accountKey) throw expired();
  return grant.bearer;
}
export async function clearPendingAppSession(env, nonce) { await env.AUTH_STATE_BUCKET.delete(await pendingKey(nonce)); }

export async function bindAppWebSession(env, origin, cookie, record, bearer) {
  if (!validGSSBearer(bearer) || record.deletedAt || record.disabledAt || !/^[a-f0-9]{16}$/.test(record.appleAccountKey || '')) throw expired();
  await write(env, await sessionKey(cookie), { origin, userId: record.userId, accountKey: record.appleAccountKey,
    bearer, issuedAt: now() });
}
export async function readAppWebSession(request, env) {
  // Never accept a caller-supplied user or trust the grant without checking the
  // current account generation, disabled/deleted state and signed web cookie.
  const user = await getUser(request, env);
  if (!user || user.authMethod !== 'apple') throw expired();
  const grant = await read(env, await sessionKey(request.headers.get('cookie')), new URL(request.url).origin, 3600);
  if (grant.userId !== user.userId || grant.accountKey !== user.account.appleAccountKey) throw expired();
  return { user, bearer: grant.bearer };
}
export async function clearAppWebSession(request, env) {
  let key;
  try { key = await sessionKey(request.headers.get('cookie')); } catch { return; }
  if (env.AUTH_STATE_BUCKET) await env.AUTH_STATE_BUCKET.delete(key);
}
