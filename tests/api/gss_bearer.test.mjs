import test from 'node:test';
import assert from 'node:assert/strict';
import { validGSSBearer } from '../../dashboard/functions/_lib/app-web-session.js';
import { appBearer, verifyTeamAccess } from '../../dashboard/functions/_lib/app-access.js';

const ACCOUNT = '0123456789abcdef';
const TEAM = 'T-11111111-1111-4111-8111-111111111111';
// Reproduce Python base64.urlsafe_b64encode(...).decode(), including padding.
const wireToken = epoch => Buffer.from(`v3|${ACCOUNT}|${epoch}|1793570000|1793573600|${'0'.repeat(64)}`)
  .toString('base64').replace(/\+/g, '-').replace(/\//g, '_');

test('GSS wire tokens retain zero, one and two padding characters through native access checks', async t => {
  const tokens = [0, 10, 100].map(wireToken);
  assert.deepEqual(tokens.map(value => value.match(/=*$/)[0].length).sort(), [0, 1, 2]);
  for (const token of tokens) {
    assert.equal(validGSSBearer(token), true);
    assert.equal(appBearer(new Request('https://playbook.example/', { headers: { Authorization: `Bearer ${token}` } })), token);
    const paths = [];
    t.mock.method(globalThis, 'fetch', async (url, options) => {
      paths.push(url);
      assert.equal(options.headers.Authorization, `Bearer ${token}`);
      assert.equal(options.redirect, 'manual');
      return url.endsWith('/v2/me')
        ? Response.json({ account_key: ACCOUNT, account_status: 'enabled', migration_complete: true })
        : Response.json({ team: { team_id: TEAM, status: 'active', role: 'leader' } });
    });
    assert.equal((await verifyTeamAccess({ GSS_API_BASE: 'https://gss.example' }, token, TEAM)).isManager, true);
    assert.equal(paths.length, 2, 'syntax acceptance must still verify identity and membership');
    t.mock.restoreAll();
  }
});

test('token syntax remains bounded and rejects embedded padding, whitespace and header injection', async t => {
  t.mock.method(globalThis, 'fetch', async () => { assert.fail('malformed token reached GSS'); });
  const stem = 'synthetic-session-token';
  for (const token of [null, '', 'short', `${stem}=tail`, `${stem}===`, `${stem}\r\nX-Test: bad`,
    `${stem}\n`, `${stem}\r`, `${stem} other`, `${stem},other`, 'a'.repeat(16385), `${'a'.repeat(16384)}=`]) {
    assert.equal(validGSSBearer(token), false);
    await assert.rejects(verifyTeamAccess({ GSS_API_BASE: 'https://gss.example' }, token, TEAM), { status: 401 });
  }
});
