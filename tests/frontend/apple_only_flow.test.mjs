import assert from 'node:assert/strict';
import test from 'node:test';
import vm from 'node:vm';
import { readFileSync } from 'node:fs';

function page(file, fetch, search = '') {
  const html = readFileSync(new URL(`../../dashboard/${file}`, import.meta.url), 'utf8');
  const elements = new Map([...html.matchAll(/<[^>]+\bid="([^"]+)"[^>]*>/g)].map(([tag, id]) => [id, {
    hidden: /\bhidden\b/.test(tag), disabled: false, textContent: '', href: '', listeners: {},
    classList: { toggle() {}, add() {}, remove() {} }, focus() {},
    addEventListener(type, callback) { this.listeners[type] = callback; },
  }]));
  const redirects = [];
  const location = { search, replace: value => redirects.push(value) };
  const context = vm.createContext({ document: { getElementById: id => {
    assert.ok(elements.has(id), `Missing element ${id}`); return elements.get(id);
  } }, window: { location }, location, fetch, URLSearchParams, AbortSignal });
  const source = [...html.matchAll(/<script>([\s\S]*?)<\/script>/g)].map(m => m[1]).join('\n');
  const ready = vm.runInContext(source, context);
  return { elements, redirects, ready };
}
const response = (body, status = 200) => new Response(JSON.stringify(body), { status });
const appleUser = { userId: 'coach-id', authMethod: 'apple', displayName: 'Coach' };

test('Apple-only home offers signup and retains only allowlisted return destinations', async () => {
  for (const [search, destination] of [['', '/'], ['?signin=converter', '/converter'],
    ['?signin=editor', '/editor'], ['?signin=https://untrusted.example', '/']]) {
    const p = page('index.html', async path => path.endsWith('/config')
      ? response({ apple: true }) : response({}, 401), search);
    await p.ready;
    assert.equal(p.elements.get('authView').hidden, false);
    assert.equal(p.elements.get('hubView').hidden, true);
    assert.equal(p.elements.get('appleSignin').hidden, false);
    assert.equal(p.elements.get('appleSigninLink').href, '/api/auth/apple/start?returnTo=' + encodeURIComponent(destination));
    assert.equal(p.elements.has('authPassword'), false);
  }
});

test('Apple config failures fail closed while existing Apple sessions still open the hub', async () => {
  for (const account of [null, appleUser]) {
    const p = page('index.html', async path => path.endsWith('/config')
      ? response({ apple: false }) : response(account, account ? 200 : 401));
    await p.ready;
    assert.equal(p.elements.get('appleSignin').hidden, true);
    assert.match(p.elements.get('appleStatus').textContent, /temporarily unavailable/);
    assert.equal(p.elements.get('hubView').hidden, !account);
  }
});

test('returning Apple coaches reach the converter and legacy account responses cannot open the hub', async () => {
  const p = page('index.html', async path => response(path.endsWith('/config') ? { apple: true } : appleUser), '?signin=converter');
  await p.ready;
  assert.deepEqual(p.redirects, ['/converter']);
  const old = page('index.html', async path => response(path.endsWith('/config') ? { apple: true }
    : { ...appleUser, authMethod: 'password' }));
  await old.ready;
  assert.equal(old.elements.get('hubView').hidden, true);
  assert.match(old.elements.get('authError').textContent, /sign in with Apple/);
});

test('first-time welcome creates only an Apple account and rejects external return URLs', async () => {
  const requests = [];
  const p = page('apple-welcome.html', async (path, options) => {
    requests.push({ path, options });
    return response(path.endsWith('/pending') ? { account: null } : { returnTo: 'https://untrusted.example' });
  });
  await p.ready;
  assert.equal(p.elements.get('create').hidden, false);
  await p.elements.get('create').listeners.click();
  assert.equal(requests[1].path, '/api/auth/apple/finish');
  assert.deepEqual(JSON.parse(requests[1].options.body), { mode: 'new' });
  assert.deepEqual(p.redirects, ['/']);
});

test('expired welcome proof offers restart without an account-creation button', async () => {
  const p = page('apple-welcome.html', async () => response({ error: 'Sign-in expired.' }, 401));
  await p.ready;
  assert.equal(p.elements.get('create').hidden, true);
  assert.equal(p.elements.get('restart').hidden, false);
});
