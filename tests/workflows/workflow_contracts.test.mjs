import assert from 'node:assert/strict';
import { readFileSync } from 'node:fs';
import test from 'node:test';

const processWorkflow = readFileSync(
  new URL('../../.github/workflows/process.yml', import.meta.url),
  'utf8'
);
const deployWorkflow = readFileSync(
  new URL('../../.github/workflows/deploy.yml', import.meta.url),
  'utf8'
);

test('dispatch values reach shell commands only through environment variables', () => {
  assert.doesNotMatch(
    processWorkflow,
    /run:\s*[^\n]*\$\{\{\s*github\.event\.client_payload\.job_id\s*\}\}/
  );
  assert.match(processWorkflow, /run:\s*python pipeline\/process_job\.py "\$JOB_ID"/);
  assert.match(processWorkflow, /run:\s*python pipeline\/mark_job_failed\.py "\$JOB_ID"/);
});

test('checkout credentials are not persisted in production workflows', () => {
  for (const workflow of [processWorkflow, deployWorkflow]) {
    assert.match(workflow, /persist-credentials:\s*false/);
  }
});
