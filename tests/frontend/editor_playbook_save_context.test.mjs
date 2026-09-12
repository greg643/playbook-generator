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
  return editor.slice(start, end);
}

function loadSaveModel() {
  const calls = [];
  let retryCount = 0;
  let conflictCount = 0;
  const context = vm.createContext({
    Headers,
    console,
    clearTimeout() {},
    setTimeout() {
      retryCount += 1;
      return retryCount;
    },
    fetch: async (url, options) => {
      calls.push({ url, options });
      return {
        ok: true,
        status: 200,
        headers: new Headers(),
        async json() {
          return { updatedAt: 'server-time', revision: 'server-revision' };
        },
      };
    },
  });

  const source = `
    let doc = null;
    let docOwner = null;
    let activeUserId = null;
    let docPlaybookId = null;
    let activePlaybookId = null;
    let documentEpoch = 0;
    let saveTimer = null;
    let changeCounter = 0;
    let saveState = 'saved';
    let saveChain = Promise.resolve(false);
    let serverUpdatedAt = null;
    let serverRevision = null;
    let overwriteNext = false;
    function isPlaybookId(value) { return typeof value === 'string' && value.length > 0; }
    function updateSaveIndicator() {}
    function showAuthView() {}
    function resolveConflict() { globalThis.__conflictCount += 1; }
    ${sourceBetween('function playbookDocumentUrl(', 'function renderPlaybookManager(')}
    ${sourceBetween('function captureDocumentContext()', '// ---------- Play document management ----------')}
    ${sourceBetween('function saveDoc()', 'let conflictOpen = false;')}
    globalThis.saveModel = {
      install(ownerId, playbookId, revision, updatedAt = 'loaded-time') {
        documentEpoch += 1;
        doc = { schema: 2, defaultPlayersPerSide: 5, offense: [], defense: [] };
        docOwner = ownerId;
        activeUserId = ownerId;
        docPlaybookId = playbookId;
        activePlaybookId = playbookId;
        serverRevision = revision;
        serverUpdatedAt = updatedAt;
        overwriteNext = false;
        changeCounter += 1;
        saveState = 'unsaved';
        return doc;
      },
      setSaveGate(promise) { saveChain = promise; },
      save: saveDoc,
      makeDirty() { changeCounter += 1; saveState = 'unsaved'; },
      forceOverwrite() { overwriteNext = true; },
      state() {
        return {
          activePlaybookId,
          docPlaybookId,
          documentEpoch,
          serverRevision,
          serverUpdatedAt,
          saveState,
        };
      },
    };
  `;
  context.__conflictCount = 0;
  new vm.Script(source, { filename: 'editor-playbook-save-context.js' }).runInContext(context);
  return {
    calls,
    context,
    model: context.saveModel,
    get retryCount() { return retryCount; },
    get conflictCount() { return context.__conflictCount; },
  };
}

test('a save captures its playbook before entering the serialized queue', async () => {
  const harness = loadSaveModel();
  let release;
  const gate = new Promise(resolve => { release = resolve; });
  harness.model.install('coach-a', 'book-a', 'revision-a');
  harness.model.setSaveGate(gate);

  const queued = harness.model.save();
  harness.model.install('coach-a', 'book-b', 'revision-b');
  release(false);

  assert.equal(await queued, false);
  assert.equal(harness.calls.length, 0);
  assert.equal(harness.model.state().activePlaybookId, 'book-b');
  assert.equal(harness.model.state().serverRevision, 'revision-b');
});

test('a delayed response from one playbook cannot mutate the next playbook', async () => {
  const harness = loadSaveModel();
  let requestStarted;
  const started = new Promise(resolve => { requestStarted = resolve; });
  let finishRequest;
  const response = new Promise(resolve => { finishRequest = resolve; });
  harness.context.fetch = async (url, options) => {
    harness.calls.push({ url, options });
    requestStarted();
    return response;
  };

  harness.model.install('coach-a', 'book-a', 'revision-a');
  const savingA = harness.model.save();
  await started;
  harness.model.install('coach-a', 'book-b', 'revision-b');
  finishRequest({
    ok: false,
    status: 409,
    headers: new Headers(),
    async json() { return { error: 'conflict' }; },
  });

  assert.equal(await savingA, false);
  assert.equal(harness.calls[0].url, '/api/plays?playbookId=book-a');
  assert.equal(harness.model.state().activePlaybookId, 'book-b');
  assert.equal(harness.model.state().serverRevision, 'revision-b');
  assert.equal(harness.model.state().saveState, 'unsaved');
  assert.equal(harness.retryCount, 0);
  assert.equal(harness.conflictCount, 0);
});

test('current saves carry the document revision and explicit overwrites carry a force flag', async () => {
  const harness = loadSaveModel();
  harness.model.install('coach-a', 'book-a', 'revision-a', 'time-a');

  assert.equal(await harness.model.save(), true);
  const normalBody = JSON.parse(harness.calls[0].options.body);
  assert.equal(harness.calls[0].url, '/api/plays?playbookId=book-a');
  assert.equal(normalBody.ownerId, 'coach-a');
  assert.equal(normalBody.baseRevision, 'revision-a');
  assert.equal(normalBody.baseUpdatedAt, 'time-a');
  assert.equal(Object.hasOwn(normalBody, 'forceOverwrite'), false);
  assert.equal(harness.model.state().serverRevision, 'server-revision');

  harness.model.makeDirty();
  harness.model.forceOverwrite();
  assert.equal(await harness.model.save(), true);
  const forcedBody = JSON.parse(harness.calls[1].options.body);
  assert.equal(forcedBody.forceOverwrite, true);
  assert.equal(Object.hasOwn(forcedBody, 'baseRevision'), false);
  assert.equal(Object.hasOwn(forcedBody, 'baseUpdatedAt'), false);
});
