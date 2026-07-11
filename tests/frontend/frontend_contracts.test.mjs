import assert from 'node:assert/strict';
import { readFileSync } from 'node:fs';
import test from 'node:test';
import vm from 'node:vm';

const read = name => readFileSync(new URL(`../../dashboard/${name}`, import.meta.url), 'utf8');
const editor = read('editor.html');
const index = read('index.html');
const help = read('help.html');

function inlineScripts(html) {
  return [...html.matchAll(/<script(?:\s[^>]*)?>([\s\S]*?)<\/script>/gi)].map(match => match[1]);
}

test('inline frontend scripts compile', () => {
  for (const [name, html] of [['editor.html', editor], ['index.html', index]]) {
    const scripts = inlineScripts(html);
    assert.ok(scripts.length, `${name} should contain an inline script`);
    scripts.forEach((source, i) => assert.doesNotThrow(
      () => new vm.Script(source, { filename: `${name}#script-${i}` })
    ));
  }
});

test('polling is sequential and bounded on both generation surfaces', () => {
  for (const html of [editor, index]) {
    assert.doesNotMatch(html, /\bsetInterval\s*\(/);
    assert.match(html, /POLL_TIMEOUT_MS\s*=\s*10\s*\*\s*60\s*\*\s*1000/);
    assert.match(html, /pollTimer\s*=\s*setTimeout\(poll,\s*POLL_INTERVAL_MS\)/);
    assert.match(html, /if\s*\(!res\.ok\)/);
    assert.match(html, /\['processing', 'complete', 'error'\]\.includes\(data\.status\)/);
  }
});

test('editor conflict and recovery dialogs preserve local work and one-time secrets', () => {
  assert.match(editor, /cancelValue:\s*'keep'/);
  assert.match(editor, /key:\s*'keep',\s*label:\s*'Keep editing'/);
  assert.match(editor, /key:\s*'reload',\s*label:\s*'Discard local changes'/);
  assert.match(editor, /key:\s*'overwrite',\s*label:\s*'Overwrite server'/);
  assert.match(editor, /function showRecoveryCode[\s\S]*?dismissible:\s*false/);
  assert.match(editor, /const modalQueue\s*=\s*\[\]/);
  assert.match(editor, /beforeunload[\s\S]*?activeModal\s*&&\s*!activeModal\.dismissible/);
});

test('output gating retains checkbox preferences and resyncs membership changes', () => {
  assert.doesNotMatch(editor, /if\s*\(!ok\)\s*box\.checked\s*=\s*false/);
  assert.match(editor, /function togglePlayIncluded[\s\S]*?syncGenerateChecks\(\)/);
  assert.match(editor, /!chkOffenseCoach\.disabled\s*&&\s*chkOffenseCoach\.checked/);
});

test('drawing tools keep independent defaults', () => {
  assert.match(editor, /const toolStyles\s*=\s*\{/);
  assert.match(editor, /route:\s*\{\s*color:\s*'#FF0000',\s*dash:\s*false/);
  assert.match(editor, /line:\s*\{\s*color:\s*'#1F6E8C',\s*dash:\s*true/);
  assert.match(editor, /if\s*\(toolStyles\[t\]\)\s*style\s*=\s*\{\s*\.\.\.toolStyles\[t\]\s*\}/);
});

test('imports and permanent save failures are bounded', () => {
  assert.match(editor, /MAX_IMPORT_BYTES\s*=\s*1000000/);
  assert.match(editor, /file\.size\s*>\s*MAX_IMPORT_BYTES/);
  assert.match(editor, /err\.permanent\s*=\s*res\.status\s*>=\s*400/);
  assert.match(editor, /Save rejected/);
});

test('unsaved editor state is scoped to an immutable account ID', () => {
  assert.match(editor, /docOwner\s*=\s*null;\s*\/\/ immutable user ID/);
  assert.match(editor, /docOwner\s*===\s*userId/);
  assert.doesNotMatch(editor, /docOwner\s*===\s*email/);
  assert.match(editor, /ownerId:\s*savingOwner/);
  assert.match(editor, /activeUserId\s*!==\s*savingOwner/);
});

test('account loads and first saves are race-safe', () => {
  assert.match(editor, /let authEpoch\s*=\s*0/);
  assert.match(editor, /playbookLoadRequest\.abort\(\)/);
  assert.match(editor, /epoch\s*!==\s*authEpoch\s*\|\|\s*activeUserId\s*!==\s*userId/);
  assert.match(editor, /if\s*\(!overwriteNext\)\s*body\.baseUpdatedAt\s*=\s*serverUpdatedAt/);
});

test('authentication and account actions cannot apply stale responses', () => {
  assert.match(editor, /let authBusy\s*=\s*false/);
  assert.match(editor, /if\s*\(authBusy\)\s*return/);
  assert.match(editor, /verifyAuthenticatedAccount\(data\.userId,\s*epoch\)/);
  assert.match(editor, /menuEpoch\s*!==\s*authEpoch\s*\|\|\s*menuUserId\s*!==\s*activeUserId/);
  assert.match(editor, /setAccountRequestBusy\(true\)/);
  assert.match(editor, /requestDeletionUntilSettled/);
  assert.match(editor, /res\.status\s*===\s*423\s*&&\s*data\.deletionPending/);
  assert.match(editor, /res\.status\s*===\s*401\s*&&\s*sawPending/);
});

test('download labels are built without an HTML injection sink', () => {
  assert.doesNotMatch(index, /\.innerHTML\s*=/);
  assert.match(index, /a\.appendChild\(document\.createTextNode/);
});

test('responsive and keyboard contracts stay present', () => {
  assert.match(editor, /@media \(max-width: 820px\)/);
  assert.match(index, /id="dropZone" role="button" tabindex="0"/);
  assert.match(index, /dropZone\.addEventListener\('keydown'/);
  assert.match(editor, /role="dialog" aria-modal="true"/);
});

test('help names the Arrow and Block controls shown by the editor', () => {
  assert.match(help, /<strong>Arrow<\/strong> tool/);
  assert.match(help, /<strong>Add a block<\/strong>/);
  assert.match(help, /Arrow \/ Ball \/ Block \/ None/);
});
