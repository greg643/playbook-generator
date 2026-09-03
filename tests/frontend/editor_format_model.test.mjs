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

// Evaluate only the format/model layer, with small stubs for the two UI side
// effects used by duplicatePlay. This exercises the production functions
// themselves without executing the editor's DOM/auth/generation bootstrap.
function loadEditorModel() {
  let uuid = 0;
  const context = vm.createContext({
    crypto: { randomUUID: () => `model-test-${++uuid}` },
  });
  const modelSource = [
    sourceBetween('const MAX_OFFENSE = 64', 'const ROUTE_COLORS ='),
    sourceBetween('function normalizePlayersPerSide(', '// ---------- Coordinate helpers ----------'),
    sourceBetween('function clamp01(', 'function dist('),
    sourceBetween('function isIncluded(', 'function displayNum('),
    sourceBetween('function makePlay(', 'function firstAvailable('),
    `
      let doc = null;
      let cur = null;
      let dirtyCount = 0;
      let lastOpened = null;
      function markDirty() { dirtyCount += 1; }
      function openPlay(section, index) {
        cur = { section, index };
        lastOpened = { section, index };
      }
    `,
    sourceBetween('function duplicatePlay(', 'async function deletePlay('),
    `
      globalThis.editorFormatModel = {
        normalizeDoc(input) {
          return normalizeDoc(input);
        },
        makePlay(section, defaultPlayersPerSide) {
          doc = { defaultPlayersPerSide };
          return makePlay(section);
        },
        duplicate(input, section, index, current = null) {
          doc = input;
          cur = current;
          dirtyCount = 0;
          lastOpened = null;
          duplicatePlay(section, index);
          return { doc, dirtyCount, lastOpened };
        },
      };
    `,
  ].join('\n');

  new vm.Script(modelSource, { filename: 'editor-format-model.js' }).runInContext(context);
  return context.editorFormatModel;
}

const model = loadEditorModel();
const plain = value => JSON.parse(JSON.stringify(value));
const sortedKeys = object => Object.keys(object).sort();

const OFFENSE_SIX_KEYS = ['1', '2', '3', '4', '5', 'QB'];
const OFFENSE_FIVE_KEYS = ['1', '2', '3', 'C', 'QB'];
const DEFENSE_FIVE_KEYS = ['1', '2', '3', '4', 'N'];
const DEFENSE_SIX_KEYS = ['1', '2', '3', '4', '5', 'N'];

function chips(keys, y = 0.7) {
  return Object.fromEntries(keys.map((key, index) => [key, {
    x: Number(((index + 1) / (keys.length + 1)).toFixed(4)),
    y,
  }]));
}

test('schema-1 offense stays an exact six-player, C-less play with positions and route anchor intact', () => {
  const legacyPositions = {
    '1': { x: 0.04, y: 0.61 },
    '2': { x: 0.18, y: 0.63 },
    '3': { x: 0.36, y: 0.67 },
    '4': { x: 0.54, y: 0.69 },
    '5': { x: 0.72, y: 0.84 },
    'QB': { x: 0.88, y: 0.78 },
  };
  const normalized = plain(model.normalizeDoc({
    schema: 1,
    offense: [{
      id: 'legacy-six',
      name: 'Legacy six',
      chips: legacyPositions,
      routes: [{
        chip: '5',
        color: '#FF0000',
        dash: false,
        end: 'arrow',
        corner: 'smooth',
        points: [[0.72, 0.84], [0.75, 0.45], [0.60, 0.20]],
      }],
      lines: [],
      labels: [],
      balls: [],
    }],
    defense: [],
  }));

  assert.equal(normalized.schema, 2);
  assert.equal(normalized.defaultPlayersPerSide, 6);
  assert.equal(normalized.offense[0].playersPerSide, 6);
  assert.deepEqual(sortedKeys(normalized.offense[0].chips), [...OFFENSE_SIX_KEYS].sort());
  assert.equal(Object.hasOwn(normalized.offense[0].chips, 'C'), false);
  assert.deepEqual(normalized.offense[0].chips, legacyPositions);
  assert.equal(normalized.offense[0].routes.length, 1);
  assert.equal(normalized.offense[0].routes[0].chip, '5');
  assert.deepEqual(
    normalized.offense[0].routes[0].points,
    [[0.72, 0.84], [0.75, 0.45], [0.60, 0.20]],
  );
});

test('new 5v5 offense and defense use the exact agreed player chips', () => {
  const offense = plain(model.makePlay('offense', 5));
  const defense = plain(model.makePlay('defense', 5));

  assert.equal(offense.playersPerSide, 5);
  assert.deepEqual(sortedKeys(offense.chips), [...OFFENSE_FIVE_KEYS].sort());
  assert.deepEqual(offense.chips.C, { x: 0.5, y: 0.66 });
  assert.deepEqual(offense.chips.QB, { x: 0.5, y: 0.88 });
  assert.equal(Object.hasOwn(offense.chips, '4'), false);
  assert.equal(Object.hasOwn(offense.chips, '5'), false);

  assert.equal(defense.playersPerSide, 5);
  assert.deepEqual(sortedKeys(defense.chips), [...DEFENSE_FIVE_KEYS].sort());
  assert.deepEqual(defense.chips.N, { x: 0.5, y: 0.62 });
  assert.equal(Object.hasOwn(defense.chips, '5'), false);
});

test('unchanged 6v6 definitions still produce the established chip sets and positions', () => {
  const offense = plain(model.makePlay('offense', 6));
  const defense = plain(model.makePlay('defense', 6));

  assert.deepEqual(sortedKeys(offense.chips), [...OFFENSE_SIX_KEYS].sort());
  assert.deepEqual(offense.chips, {
    '1': { x: 0.09, y: 0.66 },
    '2': { x: 0.21, y: 0.66 },
    '3': { x: 0.33, y: 0.66 },
    '4': { x: 0.45, y: 0.66 },
    'QB': { x: 0.62, y: 0.66 },
    '5': { x: 0.10, y: 0.88 },
  });
  assert.deepEqual(sortedKeys(defense.chips), [...DEFENSE_SIX_KEYS].sort());
  assert.deepEqual(defense.chips, {
    '2': { x: 0.30, y: 0.12 },
    '3': { x: 0.68, y: 0.12 },
    '1': { x: 0.50, y: 0.38 },
    '4': { x: 0.20, y: 0.62 },
    'N': { x: 0.50, y: 0.62 },
    '5': { x: 0.80, y: 0.62 },
  });
});

test('schema-2 default and mixed per-play formats survive repeated normalization', () => {
  const fiveOffense = plain(model.makePlay('offense', 5));
  const sixOffense = plain(model.makePlay('offense', 6));
  const fiveDefense = plain(model.makePlay('defense', 5));
  const sixDefense = plain(model.makePlay('defense', 6));
  fiveOffense.name = 'Current offense';
  sixOffense.name = 'Archived offense';
  fiveDefense.name = 'Current defense';
  sixDefense.name = 'Archived defense';

  const once = plain(model.normalizeDoc({
    schema: 2,
    defaultPlayersPerSide: 5,
    offense: [fiveOffense, sixOffense],
    defense: [fiveDefense, sixDefense],
  }));
  const twice = plain(model.normalizeDoc(JSON.parse(JSON.stringify(once))));

  assert.equal(once.defaultPlayersPerSide, 5);
  assert.deepEqual(once.offense.map(play => play.playersPerSide), [5, 6]);
  assert.deepEqual(once.defense.map(play => play.playersPerSide), [5, 6]);
  assert.deepEqual(sortedKeys(once.offense[0].chips), [...OFFENSE_FIVE_KEYS].sort());
  assert.deepEqual(sortedKeys(once.offense[1].chips), [...OFFENSE_SIX_KEYS].sort());
  assert.deepEqual(sortedKeys(once.defense[0].chips), [...DEFENSE_FIVE_KEYS].sort());
  assert.deepEqual(sortedKeys(once.defense[1].chips), [...DEFENSE_SIX_KEYS].sort());
  assert.deepEqual(twice, once);
});

test('normalization removes malformed routes and routes anchored to absent players', () => {
  const normalized = plain(model.normalizeDoc({
    schema: 2,
    defaultPlayersPerSide: 5,
    offense: [{
      id: 'five-with-routes',
      name: 'Five with routes',
      playersPerSide: 5,
      chips: chips(OFFENSE_FIVE_KEYS),
      routes: [
        { chip: 'C', points: [[0.5, 0.7], [0.5, 0.3]] },
        { chip: '4', points: [[0.4, 0.7], [0.4, 0.3]] },
        { chip: 'ghost', points: [[0.6, 0.7], [0.6, 0.3]] },
        { chip: '1', points: [[0.1, 0.7]] },
        null,
      ],
      lines: [],
      labels: [],
      balls: [],
    }],
    defense: [],
  }));

  assert.equal(normalized.offense[0].routes.length, 1);
  assert.equal(normalized.offense[0].routes[0].chip, 'C');
});

test('duplicating a play preserves its format and route model while assigning a fresh identity', () => {
  const sourceDoc = plain(model.normalizeDoc({
    schema: 2,
    defaultPlayersPerSide: 6,
    offense: [{
      id: 'five-source',
      name: 'Center choice',
      num: 1,
      playersPerSide: 5,
      chips: chips(OFFENSE_FIVE_KEYS),
      routes: [{ chip: 'C', points: [[0.5, 0.7], [0.5, 0.2]] }],
      lines: [],
      labels: [],
      balls: [],
    }],
    defense: [],
  }));
  const original = plain(sourceDoc.offense[0]);
  const result = plain(model.duplicate(sourceDoc, 'offense', 0));

  assert.equal(result.doc.offense.length, 2);
  assert.deepEqual(result.doc.offense[0], original);
  const copy = result.doc.offense[1];
  assert.equal(copy.playersPerSide, 5);
  assert.deepEqual(copy.chips, original.chips);
  assert.deepEqual(copy.routes, original.routes);
  assert.notEqual(copy.id, original.id);
  assert.equal(copy.name, 'Center choice copy');
  assert.equal(copy.num, 2);
  assert.equal(result.dirtyCount, 1);
  assert.deepEqual(result.lastOpened, { section: 'offense', index: 1 });
});
