import assert from 'node:assert/strict';
import { readFileSync } from 'node:fs';
import test from 'node:test';
import vm from 'node:vm';

const editor = readFileSync(new URL('../../dashboard/editor.html', import.meta.url), 'utf8');
const start = editor.indexOf('function publicationResult(');
const end = editor.indexOf('function startPolling(', start);
assert.ok(start >= 0 && end > start);
const result = vm.runInNewContext(`
  const JOB_ID_RE = /^[0-9a-f]{8}-[0-9a-f]{4}-[1-8][0-9a-f]{3}-[89ab][0-9a-f]{3}-[0-9a-f]{12}$/i;
  ${editor.slice(start, end)}
  publicationResult;
`);
const job = '33333333-3333-4333-8333-333333333333';
const book = '44444444-4444-4444-8444-444444444444';

test('render completion alone is not Team publication success', () => {
  assert.equal(result({ status: 'complete', files: ['offense_coach_card.pdf'] }, job), 'pending');
});
test('a confirmed current snapshot and valid book can report publication success', () => {
  assert.equal(result({ publication: { published: true, snapshotId: job, bookId: book } }, job), 'published');
});
test('a lost or negative publication receipt remains unconfirmed', () => {
  for (const publication of [{}, { published: false }, { status: 'unconfirmed' }, { published: 'true', snapshotId: job, bookId: book }]) {
    assert.equal(result({ publication }, job), 'unconfirmed');
  }
});
test('a receipt for another job cannot report this publication as successful', () => {
  assert.equal(result({ publication: { published: true, snapshotId: book, bookId: book } }, job), 'unconfirmed');
});
test('malformed or missing book identifiers cannot advertise app availability', () => {
  for (const bookId of [undefined, '', 'default', '../other', '<script>', 42]) {
    assert.equal(result({ publication: { published: true, snapshotId: job, bookId } }, job), 'unconfirmed');
  }
});
