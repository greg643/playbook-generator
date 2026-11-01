import test from 'node:test';
import assert from 'node:assert/strict';
import { beginAppPublication, completeAppPublication, readTeamPublications, snapshotPrefix } from '../../dashboard/functions/_lib/app-publications.js';
import { verifyTeamAccess } from '../../dashboard/functions/_lib/app-access.js';
import { onRequestGet as readApp } from '../../dashboard/functions/api/app/playbooks/[[path]].js';

const UID = '11111111-1111-4111-8111-111111111111';
const TEAM = 'T-22222222-2222-4222-8222-222222222222';
const OTHER = 'T-33333333-3333-4333-8333-333333333333';
const REV = '44444444-4444-4444-8444-444444444444';
const ACCOUNT = '0123456789abcdef', TOKEN = 'test-token-not-a-real-credential';
const PNG = Uint8Array.from(Buffer.from('iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVQIHWP4z8DwHwAFgAI/ScLbtAAAAABJRU5ErkJggg==', 'base64'));
const PDF = new TextEncoder().encode('%PDF-1.4\nsynthetic contract fixture\n%%EOF\n');
class R2 {
  records = new Map(); sequence = 0; beforePut = null;
  object(key, value) {
    if (!value) return null;
    const bytes = value.bytes.slice();
    return { key, size: bytes.length, etag: value.etag, checksums: { sha256: value.hash.slice(0) },
      httpMetadata: value.httpMetadata, body: new Blob([bytes]).stream(), json: async () => JSON.parse(new TextDecoder().decode(bytes)) };
  }
  async get(key) { return this.object(key, this.records.get(key)); }
  async head(key) { const object = await this.get(key); if (object) { await object.body.cancel(); delete object.body; } return object; }
  async put(key, input, options = {}) {
    if (this.beforePut) await this.beforePut(key, options);
    const bytes = typeof input === 'string' ? new TextEncoder().encode(input) : new Uint8Array(input);
    const hash = await crypto.subtle.digest('SHA-256', bytes);
    if (options.sha256) assert.equal(Buffer.from(hash).toString('hex'), options.sha256);
    // The precondition and commit are atomic in R2: do not insert an await
    // between them or this fixture invents a lost-update race in the service.
    const old = this.records.get(key);
    if ((options.onlyIf?.etagDoesNotMatch === '*' && old) ||
      (options.onlyIf?.etagMatches !== undefined && old?.etag !== options.onlyIf.etagMatches)) return null;
    const record = { bytes, hash, etag: String(++this.sequence), httpMetadata: options.httpMetadata };
    this.records.set(key, record); return this.object(key, record);
  }
  async delete(key) { this.records.delete(key); }
}
const json = (bucket, key, value) => bucket.put(key, JSON.stringify(value));
async function fixture(t) {
  const env = { PLAYBOOK_BUCKET: new R2(), JOBS_BUCKET: new R2(), GSS_API_BASE: 'https://gss.example' };
  const credentialKey = `users/byapple/${ACCOUNT}.json`;
  const account = { userId: UID, appleAccountKey: ACCOUNT, email: '', sessionVersion: 1 };
  const user = { userId: UID, email: '', credentialKey, account };
  await json(env.PLAYBOOK_BUCKET, credentialKey, account);
  await json(env.PLAYBOOK_BUCKET, `accounts/${UID}/playbooks/catalog.json`, { schema: 1,
    updatedAt: new Date().toISOString(), entries: [{ id: 'default', name: 'Black plays', revision: REV,
      createdAt: new Date().toISOString(), updatedAt: new Date().toISOString() }] });
  const document = { schema: 2, revision: REV, defaultPlayersPerSide: 5,
    offense: [{ id: 'cross', name: 'Cross', num: 9, playersPerSide: 5 }],
    defense: [{ id: 'zone', name: 'Zone', num: 3, playersPerSide: 5 }] };
  await json(env.PLAYBOOK_BUCKET, `accounts/${UID}/playbook.json`, document);
  const upstream = { role: 'leader', accountStatus: 'enabled', migration: true, forbidden: false, requests: [] };
  t.mock.method(globalThis, 'fetch', async (url, options) => {
    upstream.requests.push({ url, options });
    assert.equal(options.redirect, 'manual');
    assert.equal(options.headers.Authorization, `Bearer ${TOKEN}`);
    if (url.endsWith('/v2/me')) return Response.json({ account_key: ACCOUNT,
      account_status: upstream.accountStatus, migration_complete: upstream.migration });
    assert.ok(url.endsWith('/v2/teams/get'));
    if (upstream.forbidden || JSON.parse(options.body).team_id !== TEAM) return Response.json({}, { status: 403 });
    return Response.json({ team: { team_id: TEAM, role: upstream.role, status: 'active' } });
  });
  const input = { teamId: TEAM, operationId: crypto.randomUUID(), sourceBookId: 'default', sourceRevision: REV,
    coachesOnly: false, makeDefault: true };
  return { env, user, input, upstream, document };
}
async function render(f, input = f.input) {
  const prefix = `jobs/${input.operationId}/`;
  await json(f.env.JOBS_BUCKET, prefix + 'owner.json', { ownerId: UID });
  await json(f.env.JOBS_BUCKET, prefix + 'status.json', { status: 'complete', mode: 'images',
    publicationId: input.operationId, files: ['offense_coach_card.pdf', 'defense_coach_card.pdf'] });
  for (const file of ['plays/01.png', 'plays/D1.png']) await f.env.JOBS_BUCKET.put(prefix + file, PNG);
  for (const file of ['offense_coach_card.pdf', 'defense_coach_card.pdf']) await f.env.JOBS_BUCKET.put(prefix + file, PDF);
}
async function publish(f, input = f.input) {
  await beginAppPublication(f.env, f.user, input, TOKEN); await render(f, input);
  return completeAppPublication(f.env, f.user, input.operationId, TOKEN);
}
function read(f, path = [TEAM]) {
  return readApp({ env: f.env, params: { path }, request: new Request(`https://playbook.example/api/app/playbooks/${path.join('/')}`,
    { headers: { Authorization: `Bearer ${TOKEN}` } }) });
}

test('publishes included assets immutably and parent gets a Team default plus private download', async t => {
  const f = await fixture(t); const manifest = await publish(f);
  assert.deepEqual(manifest.plays.map(p => p.displayNumber), ['9', 'C']);
  f.upstream.role = 'member';
  const response = await read(f), library = await response.json();
  assert.equal(response.status, 200); assert.equal(library.canPublish, false);
  assert.equal(library.defaultBookId, manifest.bookId); assert.deepEqual(library.books, [manifest]);
  assert.equal(JSON.stringify(library).includes(UID), false);
  const image = await read(f, [TEAM, manifest.bookId, manifest.snapshotId, 'image-cross']);
  assert.equal(image.status, 200); assert.equal(image.headers.get('content-type'), 'image/png');
  assert.match(image.headers.get('cache-control'), /private, no-store/);
  assert.deepEqual(new Uint8Array(await image.arrayBuffer()), PNG);
});

test('missing PDF leaves the previous snapshot and default unchanged', async t => {
  const f = await fixture(t), old = await publish(f), next = { ...f.input, operationId: crypto.randomUUID() };
  await beginAppPublication(f.env, f.user, next, TOKEN); await render(f, next);
  await f.env.JOBS_BUCKET.delete(`jobs/${next.operationId}/defense_coach_card.pdf`);
  await assert.rejects(completeAppPublication(f.env, f.user, next.operationId, TOKEN));
  assert.equal((await readTeamPublications(f.env, TEAM)).value.entries[0].current.snapshotId, old.snapshotId);
  assert.equal((await (await read(f)).json()).books[0].snapshotId, old.snapshotId);
});

test('superseded render and repeated old begin cannot replace newer publication', async t => {
  const f = await fixture(t); await publish(f);
  const second = { ...f.input, operationId: crypto.randomUUID() }, third = { ...second, operationId: crypto.randomUUID() };
  await beginAppPublication(f.env, f.user, second, TOKEN); await render(f, second);
  await publish(f, third);
  await assert.rejects(completeAppPublication(f.env, f.user, second.operationId, TOKEN), { status: 409 });
  await assert.rejects(beginAppPublication(f.env, f.user, second, TOKEN), { status: 409 });
  assert.equal((await readTeamPublications(f.env, TEAM)).value.entries[0].current.snapshotId, third.operationId);
});

test('idempotent begin and completion survive lost ACK and job expiry', async t => {
  const f = await fixture(t), first = await beginAppPublication(f.env, f.user, f.input, TOKEN);
  assert.deepEqual(await beginAppPublication(f.env, f.user, f.input, TOKEN), first);
  await render(f); const manifest = await completeAppPublication(f.env, f.user, f.input.operationId, TOKEN);
  f.env.JOBS_BUCKET.records.clear();
  assert.deepEqual(await completeAppPublication(f.env, f.user, f.input.operationId, TOKEN), manifest);
  await assert.rejects(beginAppPublication(f.env, f.user, { ...f.input, coachesOnly: true }, TOKEN), { status: 409 });
});

test('editing source while a render runs cannot publish a mixed revision', async t => {
  const f = await fixture(t); await beginAppPublication(f.env, f.user, f.input, TOKEN); await render(f);
  await json(f.env.PLAYBOOK_BUCKET, `accounts/${UID}/playbook.json`, { ...f.document, revision: crypto.randomUUID() });
  await assert.rejects(completeAppPublication(f.env, f.user, f.input.operationId, TOKEN), { status: 409 });
  assert.equal((await readTeamPublications(f.env, TEAM)).value.entries[0].current, null);
});

test('revocation during asset copy prevents the pointer switch', async t => {
  const f = await fixture(t); await beginAppPublication(f.env, f.user, f.input, TOKEN); await render(f);
  f.env.PLAYBOOK_BUCKET.beforePut = async key => { if (key.endsWith('/manifest.json')) f.upstream.forbidden = true; };
  await assert.rejects(completeAppPublication(f.env, f.user, f.input.operationId, TOKEN), { status: 404 });
  assert.equal((await readTeamPublications(f.env, TEAM)).value.entries[0].current, null);
});

test('coaches-only book is absent and cannot be fetched by parents', async t => {
  const f = await fixture(t), manifest = await publish(f, { ...f.input, coachesOnly: true });
  f.upstream.role = 'member';
  const library = await (await read(f)).json();
  assert.deepEqual(library.books, []); assert.equal(library.defaultBookId, null);
  assert.equal((await read(f, [TEAM, manifest.bookId, manifest.snapshotId, 'image-cross'])).status, 404);
  f.upstream.role = 'leader'; assert.equal((await (await read(f)).json()).books.length, 1);
});

test('foreign Team access cannot enumerate books or fetch a guessed asset', async t => {
  const f = await fixture(t), manifest = await publish(f);
  assert.equal((await read(f, [OTHER])).status, 404);
  assert.equal((await read(f, [OTHER, manifest.bookId, manifest.snapshotId, 'image-cross'])).status, 404);
});

test('disabled/deleted publisher and deleted source book stop all fresh reads', async t => {
  for (const mutation of ['disabled', 'accountDeleted', 'bookDeleted']) {
    const f = await fixture(t), manifest = await publish(f);
    if (mutation === 'disabled') await json(f.env.PLAYBOOK_BUCKET, f.user.credentialKey, { ...f.user.account, disabledAt: 1 });
    else if (mutation === 'accountDeleted') await f.env.PLAYBOOK_BUCKET.delete(f.user.credentialKey);
    else await f.env.PLAYBOOK_BUCKET.delete(`accounts/${UID}/playbook.json`);
    assert.deepEqual((await (await read(f)).json()).books, []);
    assert.equal((await read(f, [TEAM, manifest.bookId, manifest.snapshotId, 'image-cross'])).status, 404);
  }
});

test('manager role, GSS identity binding and separate render bucket are required', async t => {
  const f = await fixture(t);
  f.upstream.role = 'member'; await assert.rejects(beginAppPublication(f.env, f.user, f.input, TOKEN), { status: 404 });
  f.upstream.role = 'leader';
  await assert.rejects(beginAppPublication(f.env, { ...f.user, account: { appleAccountKey: 'ffffffffffffffff' } }, f.input, TOKEN), { status: 404 });
  f.env.JOBS_BUCKET = f.env.PLAYBOOK_BUCKET;
  await assert.rejects(beginAppPublication(f.env, f.user, f.input, TOKEN), { status: 503 });
});

test('pending or unmigrated GSS identity cannot read Team books', async t => {
  const f = await fixture(t);
  f.upstream.accountStatus = 'pending'; assert.equal((await read(f)).status, 403);
  f.upstream.accountStatus = 'enabled'; f.upstream.migration = false; assert.equal((await read(f)).status, 403);
});

test('renderer cannot substitute another owner or a cancelled job', async t => {
  const f = await fixture(t); await beginAppPublication(f.env, f.user, f.input, TOKEN); await render(f);
  await json(f.env.JOBS_BUCKET, `jobs/${f.input.operationId}/owner.json`, { ownerId: crypto.randomUUID() });
  await assert.rejects(completeAppPublication(f.env, f.user, f.input.operationId, TOKEN), { status: 409 });
  await render(f); await json(f.env.JOBS_BUCKET, `jobs/${f.input.operationId}/cancelled.json`, {});
  await assert.rejects(completeAppPublication(f.env, f.user, f.input.operationId, TOKEN), { status: 409 });
});

test('immutable asset collision fails rather than replacing content', async t => {
  const f = await fixture(t), intent = await beginAppPublication(f.env, f.user, f.input, TOKEN); await render(f);
  const key = snapshotPrefix(UID, intent.bookId, intent.operationId) + 'image-cross';
  await f.env.PLAYBOOK_BUCKET.put(key, new Uint8Array([1, 2, 3]));
  await assert.rejects(completeAppPublication(f.env, f.user, f.input.operationId, TOKEN), { status: 409 });
  assert.equal((await f.env.PLAYBOOK_BUCKET.head(key)).size, 3);
});

test('two published books retain source identity and star changes default atomically', async t => {
  const f = await fixture(t), first = await publish(f), secondId = crypto.randomUUID();
  const key = `accounts/${UID}/playbooks/catalog.json`, c = await (await f.env.PLAYBOOK_BUCKET.get(key)).json();
  c.entries.push({ ...c.entries[0], id: secondId, name: 'Second book' }); await json(f.env.PLAYBOOK_BUCKET, key, c);
  await json(f.env.PLAYBOOK_BUCKET, `accounts/${UID}/playbooks/items/${secondId}.json`, f.document);
  const second = await publish(f, { ...f.input, operationId: crypto.randomUUID(), sourceBookId: secondId, makeDefault: false });
  assert.equal((await (await read(f)).json()).defaultBookId, first.bookId);
  const updated = await publish(f, { ...f.input, operationId: crypto.randomUUID(), sourceBookId: secondId, makeDefault: true });
  const library = await (await read(f)).json();
  assert.equal(library.books.length, 2); assert.equal(library.defaultBookId, second.bookId); assert.equal(updated.bookId, second.bookId);
});

test('upstream redirects and malformed/oversized responses fail closed without forwarding token', async t => {
  const f = await fixture(t);
  for (const response of [new Response(null, { status: 302, headers: { location: 'https://evil.example' } }),
    new Response('not json'), new Response('x'.repeat(33000))]) {
    t.mock.method(globalThis, 'fetch', async () => response);
    await assert.rejects(verifyTeamAccess(f.env, TOKEN, TEAM), { status: 503 });
  }
});

test('competing first publications cannot create two identities for one source book', async t => {
  const f = await fixture(t), second = { ...f.input, operationId: crypto.randomUUID() };
  const outcomes = await Promise.allSettled([
    beginAppPublication(f.env, f.user, f.input, TOKEN), beginAppPublication(f.env, f.user, second, TOKEN),
  ]);
  assert.equal(outcomes.filter(r => r.status === 'fulfilled').length, 1);
  assert.equal(outcomes.filter(r => r.status === 'rejected').length, 1);
  assert.equal((await readTeamPublications(f.env, TEAM)).value.entries.length, 1);
});

test('catalog CAS retries merge a concurrent update without losing another book', async t => {
  const f = await fixture(t); const first = await publish(f), secondId = crypto.randomUUID();
  const key = `accounts/${UID}/playbooks/catalog.json`, c = await (await f.env.PLAYBOOK_BUCKET.get(key)).json();
  c.entries.push({ ...c.entries[0], id: secondId, name: 'Another book' }); await json(f.env.PLAYBOOK_BUCKET, key, c);
  await json(f.env.PLAYBOOK_BUCKET, `accounts/${UID}/playbooks/items/${secondId}.json`, f.document);
  const next = { ...f.input, operationId: crypto.randomUUID(), makeDefault: false };
  await beginAppPublication(f.env, f.user, next, TOKEN); await render(f, next);
  let fired = false;
  f.env.PLAYBOOK_BUCKET.beforePut = async key => {
    if (key !== `teams/${TEAM}/playbooks.json` || fired) return;
    fired = true; f.env.PLAYBOOK_BUCKET.beforePut = null;
    await publish(f, { ...f.input, operationId: crypto.randomUUID(), sourceBookId: secondId, makeDefault: true });
  };
  await completeAppPublication(f.env, f.user, next.operationId, TOKEN);
  const catalog = (await readTeamPublications(f.env, TEAM)).value;
  assert.equal(catalog.entries.length, 2); assert.notEqual(catalog.defaultBookId, first.bookId);
  assert.equal(catalog.entries.find(e => e.bookId === first.bookId).current.snapshotId, next.operationId);
});

test('making a book coaches-only retires previous public snapshot asset access', async t => {
  const f = await fixture(t), old = await publish(f);
  await publish(f, { ...f.input, operationId: crypto.randomUUID(), coachesOnly: true });
  f.upstream.role = 'member';
  assert.equal((await read(f, [TEAM, old.bookId, old.snapshotId, 'image-cross'])).status, 404);
  assert.deepEqual((await (await read(f)).json()).books, []);
});

test('corrupt stored asset is not served as a valid download', async t => {
  const f = await fixture(t), m = await publish(f);
  await f.env.PLAYBOOK_BUCKET.put(snapshotPrefix(UID, m.bookId, m.snapshotId) + 'image-cross', new Uint8Array([1]));
  const response = await read(f, [TEAM, m.bookId, m.snapshotId, 'image-cross']);
  assert.equal(response.status, 503);
});

test('empty library is available to a parent without an editor account; no token gives 401', async t => {
  const f = await fixture(t); f.upstream.role = 'member';
  await f.env.PLAYBOOK_BUCKET.delete(f.user.credentialKey);
  assert.deepEqual((await (await read(f)).json()).books, []);
  const response = await readApp({ env: f.env, params: { path: [TEAM] }, request: new Request('https://playbook.example/api/app/playbooks/' + TEAM) });
  assert.equal(response.status, 401);
});

test('invalid PNG dimensions cannot produce a current snapshot', async t => {
  const f = await fixture(t); await beginAppPublication(f.env, f.user, f.input, TOKEN); await render(f);
  const huge = PNG.slice(); new DataView(huge.buffer).setUint32(16, 99999);
  await f.env.JOBS_BUCKET.put(`jobs/${f.input.operationId}/plays/01.png`, huge);
  await assert.rejects(completeAppPublication(f.env, f.user, f.input.operationId, TOKEN));
  assert.equal((await readTeamPublications(f.env, TEAM)).value.entries[0].current, null);
});

test('replaced canonical Apple identity cannot publish through a still-active legacy account', async t => {
  const f = await fixture(t), legacyKey = `users/byemail/${'a'.repeat(64)}.json`;
  await json(f.env.PLAYBOOK_BUCKET, legacyKey, f.user.account);
  f.user.credentialKey = legacyKey;
  await json(f.env.PLAYBOOK_BUCKET, `users/byapple/${ACCOUNT}.json`, {
    kind: 'link', userId: UID, credentialKey: legacyKey,
  });
  await beginAppPublication(f.env, f.user, f.input, TOKEN); await render(f);
  await json(f.env.PLAYBOOK_BUCKET, `users/byapple/${ACCOUNT}.json`, {
    userId: crypto.randomUUID(), appleAccountKey: ACCOUNT, email: '', sessionVersion: 1,
  });
  await assert.rejects(completeAppPublication(f.env, f.user, f.input.operationId, TOKEN), { status: 401 });
  assert.equal((await readTeamPublications(f.env, TEAM)).value.entries[0].current, null);
});
