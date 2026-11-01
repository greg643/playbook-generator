// The renderer holds only a one-job callback capability. The GSS session stays
// encrypted in private account storage and is revalidated at publication time.
import { getSecret, readBoundedUtf8Text } from './auth.js';
import { createJson } from './r2.js';
import { AppAccessError, isUUID } from './app-access.js';
import { beginAppPublication, completeAppPublication } from './app-publications.js';

const SCHEMA = 'gss.playbook-publication-grant.v1';
const TTL = 3600;
const bytes = value => new TextEncoder().encode(value);
const hex = value => [...new Uint8Array(value)].map(n => n.toString(16).padStart(2, '0')).join('');
const unhex = value => Uint8Array.from(value.match(/../g), n => parseInt(n, 16));
const denied = () => new AppAccessError(404, 'Publication not found');
const key = (ownerId, operationId) => {
  if (![ownerId, operationId].every(isUUID)) throw denied();
  return `accounts/${ownerId}/playbooks/publications/${operationId}.grant.json`;
};
const binding = grant => bytes([SCHEMA, grant.ownerId, grant.operationId, grant.expiresAt].join(':'));
async function cryptoKeys(env) {
  const secret = await getSecret(env);
  const aesMaterial = await crypto.subtle.digest('SHA-256', bytes(SCHEMA + ':aes:' + secret));
  const macMaterial = await crypto.subtle.digest('SHA-256', bytes(SCHEMA + ':hmac:' + secret));
  return {
    aes: await crypto.subtle.importKey('raw', aesMaterial, 'AES-GCM', false, ['encrypt', 'decrypt']),
    hmac: await crypto.subtle.importKey('raw', macMaterial, { name: 'HMAC', hash: 'SHA-256' }, false, ['sign', 'verify']),
  };
}
async function readGrant(env, ownerId, operationId) {
  const object = await env.PLAYBOOK_BUCKET.get(key(ownerId, operationId));
  if (!object) return null;
  const text = await readBoundedUtf8Text(new Response(object.body), 65536);
  let grant;
  try { grant = text && JSON.parse(text); } catch { throw denied(); }
  if (!grant || grant.schema !== SCHEMA || grant.ownerId !== ownerId || grant.operationId !== operationId ||
      !Number.isSafeInteger(grant.expiresAt) || !/^[a-f0-9]{24}$/.test(grant.iv || '') ||
      !/^(?:[a-f0-9]{2}){16,24000}$/.test(grant.sealed || '')) throw denied();
  return grant;
}

export async function preparePublicationJob(env, user, input, bearer, now = Math.floor(Date.now()/1000)) {
  const intent = await beginAppPublication(env, user, input, bearer);
  const keys = await cryptoKeys(env);
  let grant = await readGrant(env, user.userId, input.operationId);
  if (!grant) {
    grant = { schema: SCHEMA, ownerId: user.userId, operationId: input.operationId, expiresAt: now + TTL };
    const iv = crypto.getRandomValues(new Uint8Array(12));
    const payload = { bearer, user: { userId: user.userId, email: user.email, credentialKey: user.credentialKey,
      account: { appleAccountKey: user.account.appleAccountKey } } };
    const sealed = await crypto.subtle.encrypt({ name: 'AES-GCM', iv, additionalData: binding(grant) }, keys.aes,
      bytes(JSON.stringify(payload)));
    Object.assign(grant, { iv: hex(iv), sealed: hex(sealed) });
    if (await createJson(env, key(user.userId, input.operationId), grant) === null) {
      grant = await readGrant(env, user.userId, input.operationId);
    }
  }
  if (!grant || grant.expiresAt <= now || grant.expiresAt > now + TTL) throw denied();
  const token = hex(await crypto.subtle.sign('HMAC', keys.hmac, binding(grant)));
  return { intent, callback: { ownerId: user.userId, operationId: input.operationId, token } };
}

export async function finishPublicationJob(env, callback, now = Math.floor(Date.now()/1000)) {
  if (!callback || Object.keys(callback).sort().join(',') !== 'operationId,ownerId,token' ||
      !/^[a-f0-9]{64}$/.test(callback.token || '')) throw denied();
  const grant = await readGrant(env, callback.ownerId, callback.operationId);
  if (!grant || grant.expiresAt <= now || grant.expiresAt > now + TTL) throw denied();
  const keys = await cryptoKeys(env);
  if (!await crypto.subtle.verify('HMAC', keys.hmac, unhex(callback.token), binding(grant))) throw denied();
  let payload;
  try {
    const raw = await crypto.subtle.decrypt({ name: 'AES-GCM', iv: unhex(grant.iv), additionalData: binding(grant) },
      keys.aes, unhex(grant.sealed));
    payload = JSON.parse(new TextDecoder().decode(raw));
  } catch { throw denied(); }
  if (payload.user?.userId !== callback.ownerId) throw denied();
  // Exact owned job, complete assets, active account, fresh manager access,
  // unchanged source and latest pending operation are checked by the service.
  const manifest = await completeAppPublication(env, payload.user, callback.operationId, payload.bearer);
  return { published: true, bookId: manifest.bookId, snapshotId: manifest.snapshotId };
}
