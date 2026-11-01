// Local browser acceptance. Requires a static dashboard server and Playwright;
// all API traffic is synthetic and every non-local request is blocked.
// Existing Node VM checks cannot exercise native dialog focus or PNG export.
import assert from 'node:assert/strict';
import { createRequire } from 'node:module';
import { mkdir, writeFile } from 'node:fs/promises';
import { execFileSync } from 'node:child_process';
const require = createRequire(import.meta.url);
const { chromium } = require(process.env.PLAYWRIGHT_MODULE || 'playwright');
const origin = process.env.PLAYBOOK_TEST_ORIGIN || 'http://127.0.0.1:8769';
if (!['localhost', '127.0.0.1'].includes(new URL(origin).hostname)) throw Error('Local server required');
const output = process.env.UI_REVIEW_DIR || 'output/ui-review/2026-11-01-team-playbook-publish';
await mkdir(output, { recursive: true });
const browser = await chromium.launch({ headless: true, ...(process.env.CHROME_EXECUTABLE ? { executablePath: process.env.CHROME_EXECUTABLE } : {}) });
const TEAM = 'T-22222222-2222-4222-8222-222222222222', JOB = '33333333-3333-4333-8333-333333333333';
const REV = '44444444-4444-4444-8444-444444444444', SAVED = '55555555-5555-4555-8555-555555555555';
const UID = '11111111-1111-4111-8111-111111111111';
const checks = [];
const play = (id, num, name, included = true) => ({ id, num, name, included, playersPerSide: 5,
  chips: { QB: { x: .5, y: .85 }, C: { x: .5, y: .65 }, '1': { x: .15, y: .65 }, '2': { x: .3, y: .65 }, '3': { x: .8, y: .65 } },
  routes: [{ chip: '1', color: '#0070C0', end: 'arrow', points: [[.15,.65],[.15,.3],[.5,.15]] }], lines: [], labels: [], balls: [] });
const source = { schema: 2, revision: REV, defaultPlayersPerSide: 5, updatedAt: '2026-11-01T12:00:00Z',
  offense: [play('nine', 9, 'Dagger'), play('two', 2, 'Quick out'), play('bench', 4, 'Bench', false)],
  defense: [play('zone', 3, 'Cover three')] };

async function fixture({ width = 1365, height = 900, scenario = 'success', before = false } = {}) {
  const page = await browser.newPage({ viewport: { width, height }, deviceScaleFactor: 1 });
  const errors = [], uploads = [];
  let polls = 0, saveStarted = false, releaseSave = null;
  page.on('pageerror', error => errors.push(error.message));
  await page.route('**/*', async route => {
    const request = route.request(), url = new URL(request.url());
    if (url.origin !== origin) return route.abort();
    if (url.pathname === '/editor.html' && before) return route.fulfill({ contentType: 'text/html',
      body: execFileSync('git', ['show', '9df96c4:dashboard/editor.html'], { encoding: 'utf8' }) });
    if (!url.pathname.startsWith('/api/')) return route.continue();
    const json = (body, status = 200) => route.fulfill({ status, contentType: 'application/json', body: JSON.stringify(body) });
    if (url.pathname === '/api/auth/me') return json({ userId: UID, authMethod: 'apple', appleLinked: true, displayName: 'QA Coach', email: '' });
    if (url.pathname === '/api/playbooks') return json({ maxPlaybooks: 10, playbooks: [{ id: 'default', name: 'Black · Sideline book',
      revision: REV, createdAt: '2026-11-01T12:00:00Z', updatedAt: '2026-11-01T12:00:00Z' }] });
    if (url.pathname === '/api/plays' && request.method() === 'GET') return json(source);
    if (url.pathname === '/api/plays' && request.method() === 'PUT') {
      saveStarted = true;
      if (scenario === 'edit-during-save') await new Promise(resolve => { releaseSave = resolve; });
      return json({ revision: SAVED, updatedAt: '2026-11-01T12:01:00Z' });
    }
    if (url.pathname === '/api/app-publication/teams') {
      if (scenario === 'reauth') return json({ error: 'Continue with Apple again' }, 401);
      if (scenario === 'empty') return json({ teams: [], next: null });
      return json({ teams: [{ teamId: TEAM, name: 'Black', season: 'Fall 2026' }], next: null });
    }
    if (url.pathname === '/api/generate') {
      const form = await new Request(request.url(), { method: 'POST', headers: request.headers(), body: request.postDataBuffer() }).formData();
      uploads.push({ options: JSON.parse(form.get('options')), files: form.getAll('plays').map(f => ({ name: f.name, size: f.size })) });
      return json({ jobId: JOB });
    }
    if (url.pathname === '/api/status/' + JOB) {
      polls++;
      return json({ status: 'complete', files: ['offense_coach_card.pdf', 'defense_coach_card.pdf'], publicationId: JOB,
        ...(polls < 2 ? {} : { publication: scenario === 'unconfirmed' ? { published: false, status: 'unconfirmed' }
          : { published: true, bookId: JOB, snapshotId: JOB } }) });
    }
    throw Error('Unexpected synthetic API request: ' + request.method() + ' ' + url.pathname);
  });
  await page.goto(origin + '/editor.html');
  await page.locator('#editorView.active').waitFor();
  await page.locator('#playbookSelect:not([disabled])').waitFor({ timeout: 100 }).catch(() => {}); // One book correctly disables its picker.
  await page.locator('#playName:not([disabled])').waitFor();
  return { page, errors, uploads, get polls() { return polls; }, get saveStarted() { return saveStarted; }, releaseSave: () => releaseSave?.() };
}
const screenshot = (page, name) => page.screenshot({ path: output + '/' + name + '.png', fullPage: false, animations: 'disabled' });
try {
  const before = await fixture({ before: true });
  await screenshot(before.page, 'before-desktop'); await before.page.close();
  const f = await fixture();
  await f.page.locator('#publishAppShortcut').click();
  await f.page.locator('#publishConfirm:not([disabled])').waitFor();
  assert.match(await f.page.locator('#publishSummary').innerText(), /2 offense \+ 1 defense/);
  assert.equal(await f.page.locator('#publishCoachesOnly').isChecked(), false);
  assert.equal(await f.page.locator('#publishDefault').isChecked(), false);
  assert.equal(await f.page.locator('#publishTeam').inputValue(), TEAM);
  await screenshot(f.page, 'after-team-publish-desktop');
  await f.page.locator('#publishCoachesOnly').check();
  await f.page.locator('#publishDefault').check();
  assert.match(await f.page.locator('#publishAudience').innerText(), /Only this Team’s managers/);
  await screenshot(f.page, 'after-coaches-only-desktop');
  await f.page.locator('#publishConfirm').click();
  await f.page.waitForFunction(() => document.querySelector('#statusText').textContent.includes('Finishing publication'), null, { timeout: 15000 });
  assert.equal(await f.page.locator('#publicationNotice').isVisible(), false);
  assert.equal(await f.page.locator('#downloads').evaluate(el => el.classList.contains('active')), false);
  await f.page.locator('#publicationNotice:not([hidden])').waitFor({ timeout: 15000 });
  assert.match(await f.page.locator('#publicationNotice').innerText(), /Black.*Coaches only/);
  assert.equal(f.uploads.length, 1);
  assert.deepEqual(f.uploads[0].options.appPublication, { teamId: TEAM, sourceBookId: 'default', sourceRevision: SAVED, coachesOnly: true, makeDefault: true });
  assert.deepEqual(f.uploads[0].files.map(f => f.name), ['01.png', '02.png', 'D1.png']);
  assert.ok(f.uploads[0].files.every(f => f.size > 1000));
  assert.deepEqual(await f.page.evaluate(() => doc.offense.map(p => p.num)), [9,2,4]);
  await screenshot(f.page, 'after-published-desktop');
  assert.deepEqual(f.errors, []); checks.push('publish_exact_saved_revision_included_order_printed_numbers_and_confirmed_receipt');
  await f.page.close();

  for (const [name, size] of [['ipad', { width: 1024, height: 768 }], ['phone', { width: 390, height: 844 }]]) {
    const f = await fixture(size);
    await f.page.locator('#publishAppShortcut').click();
    await f.page.locator('#publishConfirm:not([disabled])').waitFor();
    const rect = await f.page.locator('#publishDialog').boundingBox();
    assert.ok(rect.x >= 0 && rect.y >= 0 && rect.x + rect.width <= size.width && rect.y + rect.height <= size.height);
    assert.equal(await f.page.locator('#publishConfirm').isVisible(), true);
    await screenshot(f.page, 'after-team-publish-' + name);
    await f.page.keyboard.press('Escape');
    assert.equal(await f.page.locator('#publishDialog').isVisible(), false);
    assert.deepEqual(f.errors, []); await f.page.close(); checks.push(name + '_dialog_fits_and_escape_returns_to_editor');
  }
  for (const scenario of ['reauth', 'empty']) {
    const f = await fixture({ scenario });
    await f.page.locator('#publishAppShortcut').click();
    await f.page.waitForFunction(() => !document.querySelector('#publishTeamStatus').textContent.includes('Loading'));
    assert.equal(await f.page.locator('#publishConfirm').isDisabled(), true);
    assert.equal(await f.page.locator('#publishReauth').isVisible(), scenario === 'reauth');
    await screenshot(f.page, 'after-' + scenario + '-desktop');
    assert.deepEqual(f.errors, []); await f.page.close(); checks.push(scenario + '_safe_publish_disabled');
  }
  const failed = await fixture({ scenario: 'unconfirmed' });
  await failed.page.locator('#publishAppShortcut').click();
  await failed.page.locator('#publishConfirm:not([disabled])').click();
  await failed.page.locator('#genError.active').waitFor({ timeout: 15000 });
  assert.match(await failed.page.locator('#genError').innerText(), /could not be confirmed/);
  assert.equal(await failed.page.locator('#publicationNotice').isVisible(), false);
  assert.equal(await failed.page.locator('#downloads').evaluate(el => el.classList.contains('active')), true);
  await screenshot(failed.page, 'after-unconfirmed-desktop');
  assert.deepEqual(failed.errors, []); await failed.page.close(); checks.push('unconfirmed_publication_preserves_pdf_downloads_without_success_claim');

  const raced = await fixture({ scenario: 'edit-during-save' });
  await raced.page.locator('#publishAppShortcut').click();
  await raced.page.locator('#publishConfirm:not([disabled])').click();
  for (let i = 0; !raced.saveStarted && i < 50; i++) await new Promise(r => setTimeout(r, 20));
  assert.equal(raced.saveStarted, true);
  await raced.page.locator('#playName').fill('Changed during save');
  raced.releaseSave();
  await raced.page.locator('#genError.active').waitFor();
  assert.match(await raced.page.locator('#genError').innerText(), /changed while saving/);
  assert.equal(raced.uploads.length, 0); assert.deepEqual(raced.errors, []);
  await raced.page.close(); checks.push('edit_during_save_prevents_stale_publication');
  await writeFile(output + '/browser-receipt.json', JSON.stringify({ status: 'PASS', runtime: 'local headless Chrome',
    data: 'synthetic fixtures; all APIs intercepted; no cloud or production calls', checks }, null, 2));
  console.log(JSON.stringify({ status: 'PASS', checks }));
} finally { await browser.close(); }
