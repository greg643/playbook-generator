import assert from 'node:assert/strict';
import { readFileSync } from 'node:fs';
import test from 'node:test';
import vm from 'node:vm';

const editor = readFileSync(
  new URL('../../dashboard/editor.html', import.meta.url),
  'utf8',
);

function sourceBetween(startMarker, endMarker) {
  const start = editor.indexOf(startMarker);
  assert.notEqual(start, -1, `missing editor source marker: ${startMarker}`);
  const end = editor.indexOf(endMarker, start + startMarker.length);
  assert.notEqual(end, -1, `missing editor source marker: ${endMarker}`);
  assert.ok(end > start, `editor source markers are out of order: ${startMarker}`);
  return editor.slice(start, end);
}

function loadPureHelpers() {
  const context = vm.createContext({});
  const source = `
    const MAX_PLAYBOOK_NAME_LENGTH = 60;
    const PLAYBOOK_NAME_CONTROL_RE = /[\\u0000-\\u001f\\u007f-\\u009f]/u;
    ${sourceBetween('function normalizePlaybookName(', 'function isPlaybookId(')}
    ${sourceBetween('function deletionSurvivorId(', 'async function reconcilePlaybookCatalogAfterDelete(')}
    this.helpers = {
      backupNameSlug,
      playbookBackupFilename,
      backupSourcePlaybookName,
      deletionSurvivorId,
    };
  `;
  new vm.Script(source, { filename: 'editor-playbook-delete-backup.js' }).runInContext(context);
  return context.helpers;
}

test('backup filenames are readable, portable, bounded, and deterministic', () => {
  const helpers = loadPureHelpers();
  assert.equal(helpers.backupNameSlug('Fall 2026 / 5v5'), 'fall-2026-5v5');
  assert.equal(helpers.backupNameSlug('Élite Ü12'), 'elite-u12');
  assert.equal(helpers.backupNameSlug('🏈🏈'), 'playbook');
  assert.equal(helpers.backupNameSlug('A'.repeat(60)).length, 48);

  const date = {
    getFullYear: () => 2026,
    getMonth: () => 8,
    getDate: () => 12,
  };
  assert.equal(
    helpers.playbookBackupFilename('Fall 2026 / 5v5', date),
    'playbook-backup-fall-2026-5v5-2026-09-12.json',
  );
});

test('backup source names are optional display metadata', () => {
  const helpers = loadPureHelpers();
  assert.equal(helpers.backupSourcePlaybookName({}), null);
  assert.equal(helpers.backupSourcePlaybookName({ playbookName: 42 }), null);
  assert.equal(helpers.backupSourcePlaybookName({ playbookName: '  Fall 2026  ' }), 'Fall 2026');
  assert.equal(helpers.backupSourcePlaybookName({ playbookName: 'bad\u0000name' }), null);

  const backupFlow = sourceBetween("$('exportBtn').addEventListener", "window.addEventListener('beforeunload'");
  assert.match(backupFlow, /schema:\s*2,[\s\S]*?playbookName:\s*current\.name/);
  assert.match(backupFlow, /const incoming = normalizeDoc\(data \|\| \{\}\)/);
  assert.match(backupFlow, /The destination name will not change\./);
  assert.doesNotMatch(backupFlow, /method:\s*'PATCH'/);
  assert.match(editor, />&#8681; Export backup<\/button>/);
  assert.match(editor, />&#8679; Import backup<\/button>/);
});

test('backup import is bound to the exact document across every await boundary', () => {
  const importFlow = sourceBetween("$('importFile').addEventListener", "window.addEventListener('beforeunload'");
  const capture = importFlow.indexOf('const importContext = captureDocumentContext()');
  const read = importFlow.indexOf('await file.text()');
  const firstGuard = importFlow.indexOf('if (!isCurrentDocumentContext(importContext))', read);
  const choice = importFlow.indexOf('await showChoiceModal', firstGuard);
  const applyGuard = importFlow.indexOf('if (!isCurrentDocumentContext(importContext))', choice);
  const replace = importFlow.indexOf("if (choice === 'replace')", applyGuard);
  assert.ok(capture >= 0 && capture < read);
  assert.ok(read >= 0 && read < firstGuard);
  assert.ok(firstGuard < choice && choice < applyGuard && applyGuard < replace);
  assert.match(importFlow, /const destination = playbookById\(importContext\.playbookId\)/);
  assert.match(importFlow, /title: sourceName[\s\S]*?Import backup from/);
});

test('delete requires an exact typed name and retires saves before the request', () => {
  const confirmation = sourceBetween(
    'function showDeletePlaybookConfirmation(',
    '// Recovery codes are queued',
  );
  assert.match(confirmation, /Export a backup first if you may want it later/);
  assert.match(confirmation, /input\.value === playbook\.name/);
  assert.match(confirmation, /input\.maxLength = playbook\.name\.length/);
  assert.match(confirmation, /confirm\.disabled = true/);
  assert.match(confirmation, /confirm\.className = 'mbtn-danger'/);

  const deletion = sourceBetween('async function deleteActivePlaybook()', 'playbookSelect.addEventListener');
  assert.match(deletion, /selected\.id === DEFAULT_PLAYBOOK_ID/);
  assert.match(deletion, /saveBlockedPlaybookId = selected\.id/);
  assert.match(deletion, /documentEpoch\+\+/);
  assert.match(deletion, /await saveChain\.catch/);
  assert.ok(deletion.indexOf('await saveChain.catch') < deletion.indexOf("method: 'DELETE'"));
  assert.match(deletion, /body: JSON\.stringify\(\{[\s\S]*?playbookId: selected\.id,[\s\S]*?baseRevision: selected\.revision/);
  assert.match(deletion, /reconcilePlaybookCatalogAfterDelete/);
  assert.match(deletion, /const outcomeIsAmbiguous = !!requestError/);
  assert.match(deletion, /openDeletionSurvivor/);
  assert.match(deletion, /if \(response && response\.status === 401\) \{\s*showAuthView/);
  assert.match(editor, /response\.status === 409[\s\S]*?data\.playbook[\s\S]*?normalizePlaybookCatalog/);
  assert.match(editor, /deletePlaybookBtn\.disabled =[\s\S]*?current\.id === DEFAULT_PLAYBOOK_ID/);
  assert.match(editor, /The original playbook cannot be deleted\./);

  const reconciliation = sourceBetween(
    'async function reconcilePlaybookCatalogAfterDelete(',
    'function retireDeletedPlaybookDocument(',
  );
  assert.match(reconciliation, /if \(retireIfUnauthorized\) retireDeletedPlaybookDocument/);

  const helpers = loadPureHelpers();
  const catalog = {
    playbooks: [{ id: 'default' }, { id: 'book-a' }, { id: 'book-c' }],
  };
  assert.equal(
    helpers.deletionSurvivorId(['default', 'book-a', 'book-b', 'book-c'], 'book-b', catalog),
    'book-c',
  );
  assert.equal(
    helpers.deletionSurvivorId(['default', 'book-a', 'book-b'], 'book-b', {
      playbooks: [{ id: 'default' }, { id: 'book-a' }],
    }),
    'book-a',
  );
});
