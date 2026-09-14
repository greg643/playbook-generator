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
const ciWorkflow = readFileSync(
  new URL('../../.github/workflows/ci.yml', import.meta.url),
  'utf8'
);
const emailWorkerConfig = JSON.parse(
  readFileSync(
    new URL('../../workers/email-sender/wrangler.jsonc', import.meta.url),
    'utf8'
  )
);

function workflowStep(workflow, name) {
  const marker = `      - name: ${name}\n`;
  const start = workflow.indexOf(marker);
  assert.notEqual(start, -1, `workflow step "${name}" is missing`);
  const nextStep = workflow.indexOf('\n      - name: ', start + marker.length);
  return workflow.slice(start, nextStep === -1 ? undefined : nextStep);
}

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

test('dashboard and email Worker changes trigger the production deployment', () => {
  assert.match(deployWorkflow, /^\s+- 'dashboard\/\*\*'$/m);
  assert.match(deployWorkflow, /^\s+- 'workers\/email-sender\/\*\*'$/m);
});

test('the private email Worker deploys before the existing Pages site', () => {
  const workerStep = workflowStep(deployWorkflow, 'Deploy private email Worker');
  const pagesStep = workflowStep(deployWorkflow, 'Deploy to Cloudflare Pages');

  assert.ok(
    deployWorkflow.indexOf(workerStep) < deployWorkflow.indexOf(pagesStep),
    'the email Worker must be available before Pages starts using its service binding'
  );
  assert.match(workerStep, /uses:\s*cloudflare\/wrangler-action@v3/);
  assert.match(workerStep, /command:\s*deploy\s*$/m);
  assert.match(workerStep, /workingDirectory:\s*workers\/email-sender\s*$/m);
  assert.match(workerStep, /apiToken:\s*\$\{\{ secrets\.CLOUDFLARE_API_TOKEN \}\}/);
  assert.match(workerStep, /accountId:\s*\$\{\{ secrets\.CLOUDFLARE_ACCOUNT_ID \}\}/);

  assert.match(pagesStep, /uses:\s*cloudflare\/wrangler-action@v3/);
  assert.match(
    pagesStep,
    /command:\s*pages deploy \. --project-name=playbook-generator --branch=main --commit-dirty=true/
  );
  assert.match(pagesStep, /workingDirectory:\s*dashboard\s*$/m);
});

test('CI installs and fully validates the email Worker', () => {
  assert.match(ciWorkflow, /^\s+email-worker:\s*$/m);
  assert.match(
    ciWorkflow,
    /working-directory:\s*workers\/email-sender\s*$/m
  );
  assert.match(
    ciWorkflow,
    /cache-dependency-path:\s*workers\/email-sender\/package-lock\.json\s*$/m
  );

  for (const command of [
    'npm ci',
    'npm run types:check',
    'npm run typecheck',
    'npm test',
    'npm run deploy:dry-run'
  ]) {
    const escapedCommand = command.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
    assert.match(ciWorkflow, new RegExp(`run:\\s*${escapedCommand}\\s*$`, 'm'));
  }
});

test('the email Worker has no public route and can only send from the GSS address', () => {
  assert.equal(emailWorkerConfig.name, 'gss-playbook-email');
  assert.equal(emailWorkerConfig.workers_dev, false);
  assert.equal(emailWorkerConfig.preview_urls, false);
  assert.equal('routes' in emailWorkerConfig, false);
  assert.deepEqual(emailWorkerConfig.send_email, [
    {
      name: 'EMAIL',
      allowed_sender_addresses: ['no-reply@greenwichsportssystems.com']
    }
  ]);
});
