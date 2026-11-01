// Publication state is private to the Pages service. Render workers get only
// JOBS_BUCKET; current pointers move once, after every permanent asset exists.
import { isAccountActive, readBoundedUtf8Text, validCredentialKey } from './auth.js';
import { findAppleAccount } from './apple.js';
import { buildAppSnapshot } from './app-snapshots.js';
import { AppAccessError, isUUID, isTeamID, verifyTeamAccess } from './app-access.js';
import { findPlaybookEntry, isPlaybookId, playbookObjectKey, readPlaybookCatalog } from './playbooks.js';
import { createJson, putJsonIfCurrent } from './r2.js';

const SCHEMA = 'gss.team-playbook-publications.v1';
const MAX_CATALOG_BYTES = 128 * 1024;
const MAX_MANIFEST_BYTES = 128 * 1024;
const conflict = () => new AppAccessError(409, 'The playbook changed. Reload and publish again.');
export const teamPublicationKey = teamId => {
  if (!isTeamID(teamId)) throw new AppAccessError(404, 'Playbook not found');
  return `teams/${teamId}/playbooks.json`;
};
export function snapshotPrefix(ownerId, bookId, snapshotId) {
  if (![ownerId, bookId, snapshotId].every(isUUID)) throw conflict();
  return `accounts/${ownerId}/playbooks/app/${bookId}/snapshots/${snapshotId}/`;
}
const intentKey = (ownerId, operationId) => {
  if (![ownerId, operationId].every(isUUID)) throw conflict();
  return `accounts/${ownerId}/playbooks/publications/${operationId}.json`;
};
const hex = buffer => [...new Uint8Array(buffer)].map(n => n.toString(16).padStart(2, '0')).join('');
const digest = async bytes => hex(await crypto.subtle.digest('SHA-256', bytes));
const same = (a, b) => JSON.stringify(a) === JSON.stringify(b);

async function jsonObject(bucket, key, limit) {
  const object = await bucket.get(key);
  if (!object) return null;
  if (!Number.isSafeInteger(object.size) || object.size > limit) { await object.body?.cancel(); throw conflict(); }
  const text = await readBoundedUtf8Text(new Response(object.body), limit);
  try { if (text === null) throw Error(); return { object, value: JSON.parse(text) }; }
  catch { throw conflict(); }
}

export async function readTeamPublications(env, teamId) {
  const result = await jsonObject(env.PLAYBOOK_BUCKET, teamPublicationKey(teamId), MAX_CATALOG_BYTES);
  if (!result) return { object: null, value: { schema: SCHEMA, teamId, defaultBookId: null, entries: [] } };
  const c = result.value;
  if (c?.schema !== SCHEMA || c.teamId !== teamId || !Array.isArray(c.entries) || c.entries.length > 100 ||
      !c.entries.every(e => e && isUUID(e.bookId) && isUUID(e.ownerId) && isPlaybookId(e.sourceBookId) && isUUID(e.lastOperation) &&
        /^[a-f0-9]{16}$/.test(e.ownerAccountKey || '') && (e.pending === null || isUUID(e.pending)) &&
        (e.current === null || (isUUID(e.current?.snapshotId) && /^[a-f0-9]{64}$/.test(e.current?.sha256 || '')))) ||
      new Set(c.entries.map(e => e.bookId)).size !== c.entries.length ||
      new Set(c.entries.map(e => `${e.ownerId}/${e.sourceBookId}`)).size !== c.entries.length ||
      (c.defaultBookId !== null && !c.entries.some(e => e.bookId === c.defaultBookId && e.current))) throw conflict();
  return result;
}

async function manager(env, user, teamId, bearer) {
  if (!user || !isUUID(user.userId) || !validCredentialKey(user.credentialKey) ||
      !(await isAccountActive(env, user))) throw new AppAccessError(401, 'Sign in again');
  const access = await verifyTeamAccess(env, bearer, teamId);
  if (access.accountKey !== user.account?.appleAccountKey || !access.isManager) {
    throw new AppAccessError(404, 'Playbook not found');
  }
  const owner = await findAppleAccount(env, access.accountKey);
  if (!owner || owner.record.userId !== user.userId || owner.credentialKey !== user.credentialKey ||
      owner.record.deletedAt || owner.record.disabledAt) throw new AppAccessError(401, 'Sign in again');
  return access;
}

async function source(env, ownerId, bookId, revision) {
  if (!isPlaybookId(bookId) || !isUUID(revision)) throw conflict();
  const { catalog } = await readPlaybookCatalog(env, ownerId);
  const entry = catalog && findPlaybookEntry(catalog, bookId);
  const loaded = await jsonObject(env.PLAYBOOK_BUCKET, playbookObjectKey(ownerId, bookId), 1_000_000);
  if (!entry || !loaded || loaded.value.revision !== revision || loaded.value.schema !== 2) throw conflict();
  return { name: entry.name, entryRevision: entry.revision, document: loaded.value };
}

export async function beginAppPublication(env, user, input, bearer) {
  if (!input || !isUUID(input.operationId) || !isPlaybookId(input.sourceBookId) ||
      !isUUID(input.sourceRevision) || typeof input.coachesOnly !== 'boolean' || typeof input.makeDefault !== 'boolean') throw conflict();
  if (!env.JOBS_BUCKET || env.JOBS_BUCKET === env.PLAYBOOK_BUCKET) {
    throw new AppAccessError(503, 'Private rendering storage is not ready');
  }
  const access = await manager(env, user, input.teamId, bearer);
  const saved = await source(env, user.userId, input.sourceBookId, input.sourceRevision);
  const selection = { teamId: input.teamId, sourceBookId: input.sourceBookId, sourceRevision: input.sourceRevision,
    coachesOnly: input.coachesOnly, makeDefault: input.makeDefault,
    layouts: { offense: input.layouts?.offense ?? 16, defense: input.layouts?.defense ?? 16 } };
  const key = intentKey(user.userId, input.operationId);
  let intent = (await jsonObject(env.PLAYBOOK_BUCKET, key, 1_100_000))?.value;
  if (intent && (!same(intent.selection, selection) || intent.ownerId !== user.userId ||
      intent.ownerAccountKey !== access.accountKey || !same(intent.source, saved))) throw conflict();
  for (let attempt = 0; attempt < 5; attempt++) {
    const { object, value: catalog } = await readTeamPublications(env, input.teamId);
    let entry = catalog.entries.find(e => e.ownerId === user.userId && e.sourceBookId === input.sourceBookId);
    if (!intent) {
      intent = { schema: SCHEMA, operationId: input.operationId, ownerId: user.userId,
        ownerAccountKey: access.accountKey, bookId: entry?.bookId || input.operationId,
        parentOperation: entry?.lastOperation || null,
        selection, source: saved, publishedAt: new Date().toISOString() };
      // Validate selection before reserving anything; placeholder descriptors
      // are discarded. Real descriptors come only from verified render bytes.
      buildAppSnapshot(snapshotInput(intent, previewDescriptors(saved.document)));
      if (await createJson(env, key, intent) === null) {
        const winner = (await jsonObject(env.PLAYBOOK_BUCKET, key, 1_100_000))?.value;
        if (!winner || !same(winner.selection, selection) || !same(winner.source, saved) || winner.bookId !== intent.bookId) throw conflict();
        intent = winner;
      }
    }
    if (entry && entry.bookId !== intent.bookId) throw conflict();
    if (entry?.current?.snapshotId === intent.operationId || entry?.pending === intent.operationId) return intent;
    // CAS-linked ordering, not wall-clock ordering: a repeated old operation or
    // delayed first request must not supersede a later publication.
    if ((entry?.lastOperation || null) !== intent.parentOperation) throw conflict();
    if (!entry) {
      if (catalog.entries.length >= 100) throw new AppAccessError(409, 'Team playbook limit reached');
      entry = { bookId: intent.bookId, ownerId: user.userId, ownerAccountKey: access.accountKey,
        sourceBookId: input.sourceBookId, current: null, pending: null, lastOperation: input.operationId };
      catalog.entries.push(entry);
    }
    entry.pending = input.operationId;
    entry.lastOperation = input.operationId;
    if (await putJsonIfCurrent(env, teamPublicationKey(input.teamId), catalog, object) !== null) return intent;
  }
  throw conflict();
}

function selectedAssets(document) {
  const entries = [];
  for (const side of ['offense', 'defense']) {
    const plays = document[side]?.filter(p => p.included !== false) || [];
    plays.forEach((p, i) => entries.push({ id: `image-${p.id}`, mediaType: 'image/png',
      file: `plays/${side === 'offense' ? String(i + 1).padStart(2, '0') : `D${i + 1}`}.png` }));
    if (plays.length) entries.push({ id: `coach-${side}`, mediaType: 'application/pdf', file: `${side}_coach_card.pdf` });
  }
  return entries;
}
function previewDescriptors(document) {
  return selectedAssets(document).map(a => ({ id: a.id, mediaType: a.mediaType, bytes: 1,
    sha256: '0'.repeat(64), ...(a.mediaType === 'image/png' ? { width: 1, height: 1 } : {}) }));
}
function snapshotInput(intent, assets) {
  return { bookId: intent.bookId, snapshotId: intent.operationId, teamId: intent.selection.teamId,
    name: intent.source.name, document: intent.source.document, expectedRevision: intent.selection.sourceRevision,
    publishedAt: intent.publishedAt, coachesOnly: intent.selection.coachesOnly, layouts: intent.selection.layouts, assets };
}

async function renderedBytes(bucket, key, asset) {
  const object = await bucket.get(key);
  const max = asset.mediaType === 'image/png' ? 4 * 1024 * 1024 : 16 * 1024 * 1024;
  if (!object || !Number.isSafeInteger(object.size) || object.size < 1 || object.size > max) {
    await object?.body?.cancel(); throw new AppAccessError(409, 'Rendered playbook is incomplete');
  }
  const bytes = new Uint8Array(await new Response(object.body).arrayBuffer());
  if (bytes.length !== object.size) throw conflict();
  const descriptor = { id: asset.id, mediaType: asset.mediaType, bytes: bytes.length, sha256: await digest(bytes) };
  if (asset.mediaType === 'image/png') {
    if (bytes.length < 33 || hex(bytes.slice(0, 8)) !== '89504e470d0a1a0a' ||
        new TextDecoder().decode(bytes.slice(12, 16)) !== 'IHDR') throw conflict();
    const view = new DataView(bytes.buffer);
    Object.assign(descriptor, { width: view.getUint32(16), height: view.getUint32(20) });
  } else if (!new TextDecoder().decode(bytes.slice(0, 8)).startsWith('%PDF-') ||
      !new TextDecoder().decode(bytes.slice(-1024)).includes('%%EOF')) throw conflict();
  return { bytes, descriptor };
}

async function immutablePut(bucket, key, bytes, mediaType, sha256) {
  const written = await bucket.put(key, bytes, { onlyIf: { etagDoesNotMatch: '*' }, sha256,
    httpMetadata: { contentType: mediaType, cacheControl: 'private, no-store' } });
  const stored = written || await bucket.head(key);
  if (!stored || stored.size !== bytes.byteLength || !stored.checksums?.sha256 || hex(stored.checksums.sha256) !== sha256) throw conflict();
}

export async function completeAppPublication(env, user, operationId, bearer) {
  const intent = (await jsonObject(env.PLAYBOOK_BUCKET, intentKey(user.userId, operationId), 1_100_000))?.value;
  if (!intent || intent.ownerId !== user.userId || intent.operationId !== operationId || intent.schema !== SCHEMA) throw conflict();
  await manager(env, user, intent.selection.teamId, bearer);
  const prior = (await readTeamPublications(env, intent.selection.teamId)).value.entries
    .find(e => e.bookId === intent.bookId && e.ownerId === user.userId);
  if (prior?.current?.snapshotId === operationId && await livePublisher(env, prior)) {
    return publicationManifest(env, intent.selection.teamId, prior);
  }
  if (!env.JOBS_BUCKET || env.JOBS_BUCKET === env.PLAYBOOK_BUCKET) throw conflict();
  const jobs = env.JOBS_BUCKET, prefix = `jobs/${operationId}/`;
  const owner = (await jsonObject(jobs, prefix + 'owner.json', 4096))?.value;
  const status = (await jsonObject(jobs, prefix + 'status.json', 8192))?.value;
  if (owner?.ownerId !== user.userId || status?.status !== 'complete' || status.mode !== 'images' ||
      status.publicationId !== operationId || !Array.isArray(status.files) ||
      await jobs.head(prefix + 'cancelled.json')) throw conflict();
  if (!same(await source(env, user.userId, intent.selection.sourceBookId, intent.selection.sourceRevision), intent.source)) throw conflict();
  const destination = snapshotPrefix(user.userId, intent.bookId, operationId), assets = [];
  // Copy one bounded object at a time. Neither the runner nor the caller chooses
  // permanent keys, source URLs, another job, or account credentials.
  for (const asset of selectedAssets(intent.source.document)) {
    if (asset.mediaType === 'application/pdf' && !status.files.includes(asset.file)) throw conflict();
    const { bytes, descriptor } = await renderedBytes(jobs, prefix + asset.file, asset);
    assets.push(descriptor);
    buildAppSnapshot(snapshotInput(intent, [...assets, ...previewDescriptors(intent.source.document).slice(assets.length)]));
    await immutablePut(env.PLAYBOOK_BUCKET, destination + asset.id, bytes, asset.mediaType, descriptor.sha256);
  }
  const manifest = buildAppSnapshot(snapshotInput(intent, assets));
  const bytes = new TextEncoder().encode(JSON.stringify(manifest)), sha256 = await digest(bytes);
  if (bytes.length > MAX_MANIFEST_BYTES) throw conflict();
  await immutablePut(env.PLAYBOOK_BUCKET, destination + 'manifest.json', bytes, 'application/json', sha256);
  for (let attempt = 0; attempt < 5; attempt++) {
    await manager(env, user, intent.selection.teamId, bearer);
    if (!same(await source(env, user.userId, intent.selection.sourceBookId, intent.selection.sourceRevision), intent.source)) throw conflict();
    if (await jobs.head(prefix + 'cancelled.json')) throw conflict();
    const { object, value: catalog } = await readTeamPublications(env, intent.selection.teamId);
    const entry = catalog.entries.find(e => e.bookId === intent.bookId && e.ownerId === user.userId);
    if (entry?.current?.snapshotId === operationId && entry.current.sha256 === sha256) return manifest;
    if (entry?.pending !== operationId) throw conflict();
    entry.current = { snapshotId: operationId, sha256 }; entry.pending = null;
    if (intent.selection.makeDefault || !catalog.defaultBookId) catalog.defaultBookId = entry.bookId;
    if (await putJsonIfCurrent(env, teamPublicationKey(intent.selection.teamId), catalog, object) !== null) return manifest;
  }
  throw conflict();
}

// Online reads reject deleted/disabled publishers and deleted source books.
// A parent need not have an editor account; only the publisher does.
export async function livePublisher(env, entry) {
  const owner = await findAppleAccount(env, entry.ownerAccountKey);
  if (!owner || owner.record.userId !== entry.ownerId || owner.record.deletedAt || owner.record.disabledAt) return false;
  const { catalog } = await readPlaybookCatalog(env, entry.ownerId);
  return !!(catalog && findPlaybookEntry(catalog, entry.sourceBookId) &&
    await env.PLAYBOOK_BUCKET.head(playbookObjectKey(entry.ownerId, entry.sourceBookId)));
}

export async function publicationManifest(env, teamId, entry) {
  if (!entry.current) return null;
  const loaded = await jsonObject(env.PLAYBOOK_BUCKET,
    snapshotPrefix(entry.ownerId, entry.bookId, entry.current.snapshotId) + 'manifest.json', MAX_MANIFEST_BYTES);
  if (!loaded) throw conflict();
  const m = loaded.value;
  if (await digest(new TextEncoder().encode(JSON.stringify(m))) !== entry.current.sha256 ||
      m.schema !== 'gss.team-playbook-snapshot.v1' || m.teamId !== teamId || m.bookId !== entry.bookId ||
      m.snapshotId !== entry.current.snapshotId || typeof m.coachesOnly !== 'boolean') throw conflict();
  return m;
}
