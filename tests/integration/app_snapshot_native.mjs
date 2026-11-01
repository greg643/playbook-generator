// Explicit $0 Mac integration gate. Existing JS/Swift unit tests cannot prove
// producer-to-consumer compatibility because each constructs its own fixtures.
// No deployment, network, real account, drawing, or source-tree mutation.
import assert from 'node:assert/strict';
import { execFileSync } from 'node:child_process';
import { createHash } from 'node:crypto';
import { mkdir, mkdtemp, readFile, writeFile } from 'node:fs/promises';
import { resolve, join } from 'node:path';
import { fileURLToPath } from 'node:url';
import { buildAppSnapshot } from '../../dashboard/functions/_lib/app-snapshots.js';

assert.equal(process.platform, 'darwin', 'This explicit gate requires macOS/Xcode; it never silently skips');
assert.ok(process.env.FLAG_APP_ROOT, 'Set FLAG_APP_ROOT to the exact flag-app checkout to qualify');
const repo = fileURLToPath(new URL('../../', import.meta.url));
const app = resolve(process.env.FLAG_APP_ROOT);
const source = join(app, 'Sources/Teams/TeamPlaybook.swift');
const hash = bytes => createHash('sha256').update(bytes).digest('hex');
const started = Date.now();
await mkdir(join(repo, 'output/checks'), { recursive: true });
const out = await mkdtemp(join(repo, 'output/checks/snapshot-native-'));
const binary = join(out, 'snapshot-native');
const run = (cmd, args, cwd = repo) => execFileSync(cmd, args,
  { cwd, encoding: 'utf8', timeout: 120000, maxBuffer: 4 * 1024 * 1024 });
const checks = [];
const receipt = { status: 'RUNNING', evidence: 'local synthetic media, actual website builder and actual Swift model; not live publication or phone/offline proof',
  websiteCommit: run('git', ['rev-parse', 'HEAD']).trim(), appCommit: run('git', ['rev-parse', 'HEAD'], app).trim(),
  builderSha256: hash(await readFile(join(repo, 'dashboard/functions/_lib/app-snapshots.js'))),
  nativeSourceSha256: hash(await readFile(source)), checks };
try {
  run('xcrun', ['swiftc', source, join(repo, 'tests/integration/AppSnapshotNative.swift'), '-o', binary]);
  run(binary, ['media', out]);
  const png = await readFile(join(out, 'play.png')), pdf = await readFile(join(out, 'coach.pdf'));
  const uuid = n => `${n.repeat(8)}-${n.repeat(4)}-4${n.repeat(3)}-8${n.repeat(3)}-${n.repeat(12)}`;
  const team = `T-${uuid('1')}`;
  const asset = (id, image = true) => ({ id, mediaType: image ? 'image/png' : 'application/pdf',
    bytes: (image ? png : pdf).length, sha256: hash(image ? png : pdf), ...(image ? { width: 320, height: 160 } : {}) });
  const args = { bookId: uuid('2'), snapshotId: uuid('3'), teamId: team, name: 'Black sideline',
    expectedRevision: uuid('4'), publishedAt: '2026-11-01T12:00:00.000Z', layouts: { offense: 9 },
    document: { schema: 2, revision: uuid('4'), offense: [
      { id: 'nine', num: 9, name: 'Cross · 左', playersPerSide: 5 },
      { id: 'bench', num: 2, name: 'Excluded', playersPerSide: 5, included: false },
      { id: 'four', num: 4, name: 'Out', playersPerSide: 5 }],
    defense: [{ id: 'zone', num: 3, name: 'Zone', playersPerSide: 6 }] },
    assets: [asset('image-nine'), asset('image-four'), asset('image-zone'),
      asset('coach-offense', false), asset('coach-defense', false)] };
  const first = buildAppSnapshot(args);
  const second = buildAppSnapshot({ ...args, bookId: uuid('5'), snapshotId: uuid('6'), coachesOnly: true });
  const library = { schema: 'gss.team-playbook-library.v1', teamId: team, canPublish: true,
    defaultBookId: second.bookId, books: [first, second] };
  async function check(name, value, mode = 'accept', requestedTeam = team) {
    const file = `${name}.json`;
    await writeFile(join(out, file), JSON.stringify(value, null, 2));
    const result = run(binary, [mode, out, file, requestedTeam]).trim();
    assert.equal(result, mode === 'accept' ? 'ACCEPTED' : 'REJECTED', name);
    checks.push(name);
  }
  await check('website_two_books_actual_swift_and_media', library);
  const legacy = structuredClone(library);
  for (const book of legacy.books) for (const play of book.plays) delete play.source;
  await check('previous_editor_snapshot_without_additive_source', legacy);
  await check('foreign_team', library, 'reject', `T-${uuid('7')}`);
  for (const [name, mutate] of [
    ['unknown_schema', x => { x.books[0].schema = 'future'; }],
    ['missing_coach_card', x => { delete x.books[0].coachCards.offense; }],
    ['unknown_default', x => { x.defaultBookId = uuid('8'); }],
    ['duplicate_play', x => { x.books[0].plays.push(x.books[0].plays[0]); }],
    ['changed_media_digest', x => { x.books[0].assets[0].sha256 = '0'.repeat(64); }],
    ['wrong_decoded_dimensions', x => { x.books[0].assets[0].width = 321; }],
    ['wrong_byte_count', x => { x.books[0].assets[0].bytes++; }],
    ['unselected_extra_asset', x => { x.books[0].assets.push(asset('image-bench')); }],
  ]) {
    const value = structuredClone(library); mutate(value); await check(name, value, 'reject');
  }
  receipt.status = 'PASS';
} catch (error) { receipt.status = 'FAIL'; receipt.error = String(error); throw error; }
finally {
  receipt.seconds = (Date.now() - started) / 1000;
  await writeFile(join(out, 'receipt.json'), JSON.stringify(receipt, null, 2));
  console.log(JSON.stringify({ ...receipt, output: out }));
}
