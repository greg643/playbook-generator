import test from 'node:test';
import assert from 'node:assert/strict';
import { createSessionCookie, generatePasswordResetToken } from '../../dashboard/functions/_lib/auth.js';
import { bindAppWebSession, readAppWebSession, savePendingAppSession, readPendingAppSession,
  clearPendingAppSession } from '../../dashboard/functions/_lib/app-web-session.js';
import { onRequestPost as logout } from '../../dashboard/functions/api/auth/logout.js';
import { onRequestPost as teams } from '../../dashboard/functions/api/app-publication/teams.js';

const ORIGIN = 'https://playbook.example', ACCOUNT = '0123456789abcdef';
const UID = '11111111-1111-4111-8111-111111111111', TEAM = 'T-22222222-2222-4222-8222-222222222222';
const TOKEN = 'synthetic-gss-session-not-a-real-credential';
class Bucket {
  records = new Map();
  async put(key, text) { this.records.set(key, text); return {}; }
  async get(key) {
    const text = this.records.get(key);
    return text === undefined ? null : { body: new Blob([text]).stream(), json: async () => JSON.parse(text) };
  }
  async delete(key) { this.records.delete(key); }
}
async function fixture() {
  const env = { PLAYBOOK_BUCKET: new Bucket(), AUTH_STATE_BUCKET: new Bucket(), JOBS_BUCKET: new Bucket(),
    SESSION_SECRET: 'a'.repeat(64), GSS_API_BASE: 'https://gss.example' };
  const record = { userId: UID, appleAccountKey: ACCOUNT, email: '', sessionVersion: 1 };
  const key = `users/byapple/${ACCOUNT}.json`;
  await env.PLAYBOOK_BUCKET.put(key, JSON.stringify(record));
  const cookie = (await createSessionCookie(UID, '', env, 1, { credentialKey: key, appleAccountKey: ACCOUNT })).split(';')[0];
  await bindAppWebSession(env, ORIGIN, cookie, record, TOKEN);
  return { env, record, key, cookie };
}
const request = (cookie, body = {}, origin = ORIGIN) => new Request(ORIGIN + '/api/app-publication/teams', {
  method: 'POST', headers: { cookie, origin, 'Content-Type': 'application/json' }, body: JSON.stringify(body),
});

test('GSS bearer stays encrypted, out of cookies/account/renderer buckets, and is scoped to the exact web session', async () => {
  const f = await fixture();
  const session = await readAppWebSession(request(f.cookie), f.env);
  assert.equal(session.bearer, TOKEN); assert.equal(session.user.userId, UID);
  assert.equal(f.env.AUTH_STATE_BUCKET.records.size, 1);
  assert.equal([...f.env.AUTH_STATE_BUCKET.records.values()].join('').includes(TOKEN), false);
  assert.equal([...f.env.PLAYBOOK_BUCKET.records.values()].join('').includes(TOKEN), false);
  assert.equal(f.cookie.includes(TOKEN), false); assert.equal(f.env.JOBS_BUCKET.records.size, 0);
  await assert.rejects(readAppWebSession(request(f.cookie + 'x'), f.env), { status: 401 });
  await assert.rejects(readAppWebSession(request(f.cookie + '; ' + f.cookie), f.env), { status: 401 });
  const crossOrigin = new Request('https://other.example/', { headers: { cookie: f.cookie } });
  await assert.rejects(readAppWebSession(crossOrigin, f.env), { status: 401 });
});

test('private grant expiry leaves ordinary editor login valid but requires Apple for publishing', async t => {
  const f = await fixture(), clock = Date.now;
  t.mock.method(Date, 'now', () => clock() + 3600_000);
  await assert.rejects(readAppWebSession(request(f.cookie), f.env), { status: 401 });
  assert.equal((await teams({ env: f.env, request: request(f.cookie) })).status, 401);
});

test('account generation rotation, disable and delete fence grant reads immediately', async () => {
  for (const change of [{ sessionVersion: 2 }, { disabledAt: 'now' }, { deletedAt: 'now' }, { userId: crypto.randomUUID() }]) {
    const f = await fixture();
    await f.env.PLAYBOOK_BUCKET.put(f.key, JSON.stringify({ ...f.record, ...change }));
    await assert.rejects(readAppWebSession(request(f.cookie), f.env), { status: 401 });
  }
});

test('grant tampering and copying ciphertext to another session fail closed', async t => {
  const f = await fixture(), [key, text] = [...f.env.AUTH_STATE_BUCKET.records][0];
  const value = JSON.parse(text); value.sealed = (value.sealed[0] === '0' ? '1' : '0') + value.sealed.slice(1);
  f.env.AUTH_STATE_BUCKET.records.set(key, JSON.stringify(value));
  await assert.rejects(readAppWebSession(request(f.cookie), f.env), { status: 401 });
  const clock = Date.now;
  t.mock.method(Date, 'now', () => clock() + 1000);
  const newer = (await createSessionCookie(UID, '', f.env, 1, { credentialKey: f.key, appleAccountKey: ACCOUNT })).split(';')[0];
  await bindAppWebSession(f.env, ORIGIN, newer, f.record, TOKEN);
  const otherKey = [...f.env.AUTH_STATE_BUCKET.records.keys()].find(k => k !== key);
  f.env.AUTH_STATE_BUCKET.records.set(otherKey, text);
  await assert.rejects(readAppWebSession(request(newer), f.env), { status: 401 });
});

test('logout removes the private credential; cross-site logout cannot revoke it', async () => {
  const f = await fixture();
  assert.equal((await logout({ env: f.env, request: request(f.cookie, {}, 'https://evil.example') })).status, 403);
  assert.equal((await readAppWebSession(request(f.cookie), f.env)).bearer, TOKEN);
  const response = await logout({ env: f.env, request: request(f.cookie) });
  assert.equal(response.status, 200); assert.match(response.headers.get('set-cookie'), /Max-Age=0/);
  assert.equal(f.env.AUTH_STATE_BUCKET.records.size, 0);
  await assert.rejects(readAppWebSession(request(f.cookie), f.env), { status: 401 });
});

test('new-account private handoff is bounded, origin/account-bound, single consumed and never stored in renderer', async t => {
  const f = await fixture(), nonce = generatePasswordResetToken();
  await savePendingAppSession(f.env, ORIGIN, nonce, ACCOUNT, TOKEN);
  assert.equal(await readPendingAppSession(f.env, ORIGIN, nonce, ACCOUNT), TOKEN);
  await assert.rejects(readPendingAppSession(f.env, ORIGIN, nonce, 'fedcba9876543210'), { status: 401 });
  await assert.rejects(readPendingAppSession(f.env, 'https://evil.example', nonce, ACCOUNT), { status: 401 });
  const clock = Date.now;
  t.mock.method(Date, 'now', () => clock() + 600_000);
  await assert.rejects(readPendingAppSession(f.env, ORIGIN, nonce, ACCOUNT), { status: 401 });
  t.mock.restoreAll();
  await clearPendingAppSession(f.env, nonce);
  await assert.rejects(readPendingAppSession(f.env, ORIGIN, nonce, ACCOUNT), { status: 401 });
  assert.equal(f.env.JOBS_BUCKET.records.size, 0);
});

test('Team picker uses authenticated paginated GSS discovery, excludes members/families and leaks no credentials', async t => {
  const f = await fixture(), calls = [];
  t.mock.method(globalThis, 'fetch', async (url, options) => {
    calls.push({ url, body: JSON.parse(options.body) });
    assert.equal(options.headers.Authorization, `Bearer ${TOKEN}`); assert.equal(options.redirect, 'manual');
    if (url.endsWith('/v2/me')) return Response.json({ account_key: ACCOUNT, account_status: 'enabled', migration_complete: true });
    assert.equal(url, 'https://gss.example/v2/teams/list');
    assert.deepEqual(JSON.parse(options.body), { limit: 25, continuation: 'server-cursor' });
    return Response.json({ teams: [
      { team_id: TEAM, name: 'Black', role: 'leader', status: 'active', active_season_label: 'Fall 2026', privateData: TOKEN },
      { team_id: TEAM, name: 'Member', role: 'member', status: 'active' },
      { team_id: TEAM, name: 'Family', role: 'leader', is_family_member: true, status: 'active' },
      { team_id: TEAM, name: 'Gone', role: 'leader', status: 'inactive' },
    ], next: 'second-page' });
  });
  const response = await teams({ env: f.env, request: request(f.cookie, { continuation: 'server-cursor' }) });
  assert.equal(response.status, 200); assert.match(response.headers.get('cache-control'), /no-store/);
  assert.deepEqual(await response.json(), { teams: [{ teamId: TEAM, name: 'Black', season: 'Fall 2026' }], next: 'second-page' });
  assert.equal(calls.length, 2); assert.deepEqual(calls[0].body, { include_teams: false });
});

test('Team discovery rejects CSRF, client-supplied tokens, mismatched identity and provider redirects', async t => {
  const f = await fixture(); let count = 0, scenario = 'identity';
  t.mock.method(globalThis, 'fetch', async () => {
    count++;
    if (scenario === 'redirect') return new Response(null, { status: 302, headers: { Location: 'https://evil.example' } });
    return Response.json({ account_key: 'fedcba9876543210', account_status: 'enabled', migration_complete: true });
  });
  assert.equal((await teams({ env: f.env, request: request(f.cookie, {}, 'https://evil.example') })).status, 403);
  assert.equal((await teams({ env: f.env, request: request(f.cookie, { bearer: TOKEN }) })).status, 400);
  assert.equal(count, 0);
  assert.equal((await teams({ env: f.env, request: request(f.cookie) })).status, 401);
  scenario = 'redirect';
  assert.equal((await teams({ env: f.env, request: request(f.cookie) })).status, 503);
  assert.equal(count, 2);
});
