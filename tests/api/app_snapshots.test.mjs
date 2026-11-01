import test from 'node:test';
import assert from 'node:assert/strict';
import { buildAppSnapshot } from '../../dashboard/functions/_lib/app-snapshots.js';

const uuid = '12345678-1234-4234-8234-123456789012';
const asset = (id, mediaType = 'image/png') => ({ id, mediaType, bytes: 1024, sha256: 'a'.repeat(64),
  ...(mediaType === 'image/png' ? { width: 1280, height: 720 } : {}) });
function input() {
  return { bookId: uuid, snapshotId: uuid, teamId: `T-${uuid}`, name: 'Black sideline',
    expectedRevision: uuid, publishedAt: '2026-11-01T12:00:00.000Z',
    document: { schema: 2, revision: uuid, offense: [
      { id: 'first', num: 9, name: 'Cross', playersPerSide: 5 },
      { id: 'bench', num: 2, name: 'Do not publish', playersPerSide: 6, included: false },
      { id: 'second', num: 4, name: 'Out', playersPerSide: 5 }],
    defense: [{ id: 'defense', num: 3, name: 'Zone', playersPerSide: 6 }] },
    assets: [asset('image-first'), asset('image-second'), asset('image-defense'),
      asset('coach-offense', 'application/pdf'), asset('coach-defense', 'application/pdf')] };
}

test('snapshot preserves included order, sticky numbers, defense letters and mixed format', () => {
  const args = input(), before = structuredClone(args), result = buildAppSnapshot(args);
  assert.deepEqual(result.plays.map(p => [p.id, p.number, p.displayNumber]),
    [['first', 9, '9'], ['second', 4, '4'], ['defense', 3, 'C']]);
  assert.equal(result.format, 'mixed');
  assert.equal(result.coachesOnly, false);
  assert.equal(result.coachCards.offense.playsPerPage, 16);
  assert.deepEqual(args, before);
});

test('source rename/printed number do not change snapshot play identity', () => {
  const args = input(), first = buildAppSnapshot(args);
  args.document.offense[0].name = 'New name'; args.document.offense[0].num = 21;
  assert.equal(buildAppSnapshot(args).plays[0].id, first.plays[0].id);
});

test('offense-only book omits defense PDF and does not renumber gaps', () => {
  const args = input(); args.document.defense = [];
  args.assets = args.assets.filter(a => !a.id.includes('defense'));
  args.coachesOnly = true; args.layouts = { offense: 9 };
  const result = buildAppSnapshot(args);
  assert.equal(result.format, '5v5'); assert.equal(result.coachesOnly, true);
  assert.deepEqual(result.coachCards, { offense: { assetId: 'coach-offense', playsPerPage: 9 } });
  assert.deepEqual(result.plays.map(p => p.number), [9, 4]);
});

for (const [label, change] of [
  ['stale source', a => { a.expectedRevision = '22345678-1234-4234-8234-123456789012'; }],
  ['missing PDF', a => a.assets.pop()],
  ['extra bench image', a => a.assets.push(asset('image-bench'))],
  ['duplicate render', a => a.assets.push(a.assets[0])],
  ['duplicate number', a => { a.document.offense[2].num = 9; }],
  ['missing stable ID', a => { delete a.document.offense[0].id; }],
  ['invalid team', a => { a.teamId = '../other-team'; }],
  ['null book name', a => { a.name = null; }],
  ['bad digest', a => { a.assets[0].sha256 = 'not-hashed'; }],
  ['huge image dimensions', a => { a.assets[0].height = 99999; }],
  ['bad layout', a => { a.layouts = { offense: 7 }; }],
  ['unsupported source', a => { a.document.schema = 3; }],
  ['empty selection', a => { a.document.offense = []; a.document.defense = []; a.assets = []; }],
]) test(`publication rejects ${label}`, () => {
  const args = input(); change(args); assert.throws(() => buildAppSnapshot(args));
});

test('asset descriptors cannot leak source URLs or storage keys', () => {
  const args = input(); Object.assign(args.assets[0], { key: 'private/key', url: 'https://private.example', token: 'private' });
  const result = buildAppSnapshot(args);
  assert.deepEqual(Object.keys(result.assets[0]).sort(), ['bytes', 'height', 'id', 'mediaType', 'sha256', 'width']);
});

test('missing and numeric play IDs are rejected even when matching images exist', () => {
  for (const id of [undefined, null, 123]) {
    const args = input(); args.document.offense[0].id = id; args.assets[0].id = `image-${id}`;
    assert.throws(() => buildAppSnapshot(args), /Save valid play IDs/);
  }
});

test('render descriptors reject values that merely stringify as valid IDs or digests', () => {
  const args = input(); args.assets[0].id = { toString: () => 'image-first' };
  assert.throws(() => buildAppSnapshot(args), /Invalid render asset set/);
  const other = input(); other.assets[0].sha256 = { toString: () => 'a'.repeat(64) };
  assert.throws(() => buildAppSnapshot(other), /Missing or invalid rendered asset/);
});

test('stable play IDs leave room for the image asset prefix', () => {
  const args = input(); args.document.offense[0].id = 'x'.repeat(122);
  args.assets[0].id = `image-${args.document.offense[0].id}`;
  assert.equal(buildAppSnapshot(args).plays[0].id.length, 122);
  args.document.offense[0].id += 'x';
  assert.throws(() => buildAppSnapshot(args), /Save valid play IDs/);
});
