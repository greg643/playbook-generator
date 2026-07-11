import assert from "node:assert/strict";
import test from "node:test";

import {
  createRecoveryFields,
  createSessionCookie,
  emailKey,
  generateSaltHex,
  getUser,
  hashPassword,
} from "../../dashboard/functions/_lib/auth.js";
import {
  cancelAndScrubUserJobs,
  cleanupFailedJob,
  deleteUserQuotaRecords,
  finishJobSlot,
  listUserJobIds,
  reserveJobSlot,
} from "../../dashboard/functions/_lib/jobs.js";
import { onRequestPost as recover } from "../../dashboard/functions/api/auth/recover.js";
import { onRequestPost as deleteAccount } from "../../dashboard/functions/api/auth/delete-account.js";
import { onRequestPost as login } from "../../dashboard/functions/api/auth/login.js";
import { onRequestPost as register } from "../../dashboard/functions/api/auth/register.js";
import { onRequestPost as generate } from "../../dashboard/functions/api/generate.js";
import { onRequestPut as savePlays } from "../../dashboard/functions/api/plays.js";
import { onRequestGet as getStatus } from "../../dashboard/functions/api/status/[[jobId]].js";
import { onRequestGet as download } from "../../dashboard/functions/api/download/[[catchall]].js";

const SESSION_SECRET = "0123456789abcdef".repeat(4);
const USER_ID = "11111111-1111-4111-8111-111111111111";
const OTHER_USER_ID = "22222222-2222-4222-8222-222222222222";
const JOB_ID = "aaaaaaaa-aaaa-4aaa-8aaa-aaaaaaaaaaaa";
const TODAY = new Date().toISOString().slice(0, 10);

async function bodyBytes(value) {
  if (typeof value === "string") return new TextEncoder().encode(value);
  if (value instanceof Blob) return new Uint8Array(await value.arrayBuffer());
  if (value instanceof ArrayBuffer) return new Uint8Array(value);
  if (ArrayBuffer.isView(value)) {
    return new Uint8Array(value.buffer.slice(value.byteOffset, value.byteOffset + value.byteLength));
  }
  if (value === null) return new Uint8Array();
  throw new TypeError(`Unsupported R2 test value: ${typeof value}`);
}

class MemoryR2 {
  constructor() {
    this.objects = new Map();
    this.sequence = 0;
    this.beforeConditionalPut = null;
  }

  async put(key, value, options = {}) {
    if (options.onlyIf && this.beforeConditionalPut) {
      const hook = this.beforeConditionalPut;
      this.beforeConditionalPut = null;
      await hook(key, options);
    }

    const current = this.objects.get(key);
    const condition = options.onlyIf;
    if (condition) {
      if (condition.etagDoesNotMatch === "*" && current) return null;
      if (condition.etagMatches !== undefined && (!current || current.etag !== condition.etagMatches)) {
        return null;
      }
    }

    const bytes = await bodyBytes(value);
    const stored = {
      bytes,
      etag: `etag-${++this.sequence}`,
      httpMetadata: options.httpMetadata || {},
    };
    this.objects.set(key, stored);
    return this.object(key, stored);
  }

  object(key, stored) {
    const bytes = stored.bytes.slice();
    return {
      key,
      etag: stored.etag,
      body: bytes,
      httpMetadata: stored.httpMetadata,
      async text() {
        return new TextDecoder().decode(bytes);
      },
      async json() {
        return JSON.parse(new TextDecoder().decode(bytes));
      },
    };
  }

  async get(key) {
    const stored = this.objects.get(key);
    return stored ? this.object(key, stored) : null;
  }

  async delete(keys) {
    for (const key of Array.isArray(keys) ? keys : [keys]) this.objects.delete(key);
  }

  async list({ prefix = "", limit = 1000 } = {}) {
    const objects = [...this.objects.keys()]
      .filter((key) => key.startsWith(prefix))
      .sort()
      .slice(0, limit)
      .map((key) => ({ key }));
    return { objects, truncated: false };
  }
}

class FailFirstPlaybookDeleteR2 extends MemoryR2 {
  constructor() {
    super();
    this.failed = false;
  }

  async delete(keys) {
    const list = Array.isArray(keys) ? keys : [keys];
    if (!this.failed && list.some((key) => key.endsWith("/playbook.json"))) {
      this.failed = true;
      throw new Error("temporary storage failure");
    }
    return super.delete(keys);
  }
}

class FailFirstJobPayloadDeleteR2 extends MemoryR2 {
  constructor() {
    super();
    this.failed = false;
  }

  async delete(keys) {
    const list = Array.isArray(keys) ? keys : [keys];
    if (!this.failed && list.some((key) => key.startsWith("jobs/"))) {
      this.failed = true;
      throw new Error("temporary job-bucket failure");
    }
    return super.delete(keys);
  }
}

function makeEnv(bucket = new MemoryR2(), overrides = {}) {
  return { PLAYBOOK_BUCKET: bucket, SESSION_SECRET, ...overrides };
}

async function seedAccount(env, {
  userId = USER_ID,
  email = "coach@example.com",
  sessionVersion = 1,
  extra = {},
} = {}) {
  const record = { userId, email, sessionVersion, ...extra };
  await env.PLAYBOOK_BUCKET.put(await emailKey(email), JSON.stringify(record));
  return record;
}

async function sessionRequest(env, account, url = "https://example.test/api") {
  const setCookie = await createSessionCookie(
    account.userId,
    account.email,
    env,
    account.sessionVersion
  );
  return new Request(url, { headers: { cookie: setCookie.split(";", 1)[0] } });
}

async function json(response) {
  return response.json();
}

test("sessions are revoked by account version changes", async () => {
  const env = makeEnv();
  const account = await seedAccount(env);
  const request = await sessionRequest(env, account);

  assert.equal((await getUser(request, env)).userId, USER_ID);
  await env.PLAYBOOK_BUCKET.put(
    await emailKey(account.email),
    JSON.stringify({ ...account, sessionVersion: 2 })
  );
  assert.equal(await getUser(request, env), null);
});

test("a stale cookie cannot cross into a re-created email account", async () => {
  const env = makeEnv();
  const oldAccount = await seedAccount(env);
  const request = await sessionRequest(env, oldAccount);

  await env.PLAYBOOK_BUCKET.put(
    await emailKey(oldAccount.email),
    JSON.stringify({ ...oldAccount, userId: OTHER_USER_ID })
  );
  assert.equal(await getUser(request, env), null);
});

test("password recovery atomically consumes its code and revokes the old session", async () => {
  const env = makeEnv();
  const email = "recover@example.com";
  const salt = generateSaltHex();
  const { recoveryCode, fields } = await createRecoveryFields();
  const account = await seedAccount(env, {
    email,
    extra: {
      salt,
      iterations: 100000,
      hash: await hashPassword("old-password", salt, 100000),
      ...fields,
    },
  });
  const oldRequest = await sessionRequest(env, account);
  const recoveryBody = JSON.stringify({ email, recoveryCode, newPassword: "new-password" });

  const [first, second] = await Promise.all([
    recover({ request: new Request("https://example.test/api/auth/recover", {
      method: "POST",
      body: recoveryBody,
    }), env }),
    recover({ request: new Request("https://example.test/api/auth/recover", {
      method: "POST",
      body: recoveryBody,
    }), env }),
  ]);

  assert.deepEqual([first.status, second.status].sort((a, b) => a - b), [200, 409]);
  assert.equal(await getUser(oldRequest, env), null);
  const updated = await (await env.PLAYBOOK_BUCKET.get(await emailKey(email))).json();
  assert.equal(updated.sessionVersion, 2);
  assert.equal(updated.iterations, 600000);
});

test("job quotas limit concurrency and release terminal jobs", async () => {
  const env = makeEnv(undefined, { MAX_ACTIVE_JOBS_PER_USER: "2" });
  const now = Date.parse("2026-07-11T12:00:00Z");
  const one = "00000000-0000-4000-8000-000000000001";
  const two = "00000000-0000-4000-8000-000000000002";
  const three = "00000000-0000-4000-8000-000000000003";

  assert.equal((await reserveJobSlot(env, USER_ID, one, now)).ok, true);
  assert.equal((await reserveJobSlot(env, USER_ID, two, now)).ok, true);
  assert.equal((await reserveJobSlot(env, USER_ID, three, now)).reason, "active");
  await finishJobSlot(env, USER_ID, one, now + 1000);
  assert.equal((await reserveJobSlot(env, USER_ID, three, now + 2000)).ok, true);
});

test("job quotas enforce a configurable daily cap", async () => {
  const env = makeEnv(undefined, {
    MAX_ACTIVE_JOBS_PER_USER: "10",
    MAX_DAILY_JOBS_PER_USER: "2",
  });
  const now = Date.parse("2026-07-11T12:00:00Z");
  assert.equal((await reserveJobSlot(env, USER_ID, "00000000-0000-4000-8000-000000000011", now)).ok, true);
  assert.equal((await reserveJobSlot(env, USER_ID, "00000000-0000-4000-8000-000000000012", now)).ok, true);
  assert.equal(
    (await reserveJobSlot(env, USER_ID, "00000000-0000-4000-8000-000000000013", now)).reason,
    "daily"
  );
});

test("account cleanup removes all quota metadata without touching other users", async () => {
  const env = makeEnv();
  await env.PLAYBOOK_BUCKET.put(`accounts/${USER_ID}/job-quota/2026-07-10.json`, "{}");
  await env.PLAYBOOK_BUCKET.put(`accounts/${USER_ID}/job-quota/2026-07-11.json`, "{}");
  await env.PLAYBOOK_BUCKET.put(`accounts/${OTHER_USER_ID}/job-quota/2026-07-11.json`, "{}");
  await deleteUserQuotaRecords(env, USER_ID);
  assert.equal(
    (await env.PLAYBOOK_BUCKET.list({ prefix: `accounts/${USER_ID}/job-quota/` })).objects.length,
    0
  );
  assert.equal(
    (await env.PLAYBOOK_BUCKET.list({ prefix: `accounts/${OTHER_USER_ID}/job-quota/` })).objects.length,
    1
  );
});

test("account cleanup cancels known jobs and removes their user payload", async () => {
  const jobsBucket = new MemoryR2();
  const env = makeEnv(undefined, { JOBS_BUCKET: jobsBucket });
  await env.PLAYBOOK_BUCKET.put(
    `accounts/${USER_ID}/job-quota/${TODAY}.json`,
    JSON.stringify({
      schema: 1,
      date: TODAY,
      jobs: [{ id: JOB_ID, createdAt: `${TODAY}T12:00:00Z`, activeUntil: `${TODAY}T12:30:00Z` }],
    })
  );
  await jobsBucket.put(`jobs/${JOB_ID}/owner.json`, JSON.stringify({ ownerId: USER_ID }));
  await jobsBucket.put(`jobs/${JOB_ID}/status.json`, JSON.stringify({ status: "processing" }));
  await jobsBucket.put(`jobs/${JOB_ID}/input.pptx`, "sensitive-input");
  await jobsBucket.put(`jobs/${JOB_ID}/offense_coach_card.pdf`, "sensitive-output");

  const jobIds = await deleteUserQuotaRecords(env, USER_ID);
  assert.deepEqual(jobIds, [JOB_ID]);
  await cancelAndScrubUserJobs(env, USER_ID, jobIds, new Date("2026-07-11T12:05:00Z"));

  assert.equal(await jobsBucket.get(`jobs/${JOB_ID}/input.pptx`), null);
  assert.equal(await jobsBucket.get(`jobs/${JOB_ID}/offense_coach_card.pdf`), null);
  assert.ok(await jobsBucket.get(`jobs/${JOB_ID}/owner.json`));
  assert.ok(await jobsBucket.get(`jobs/${JOB_ID}/cancelled.json`));
  assert.equal(
    (await (await jobsBucket.get(`jobs/${JOB_ID}/status.json`)).json()).status,
    "error"
  );
});

test("ownerless reserved-job cancellation is idempotent across retries", async () => {
  const jobsBucket = new MemoryR2();
  const env = makeEnv(undefined, { JOBS_BUCKET: jobsBucket });

  assert.deepEqual(await cancelAndScrubUserJobs(env, USER_ID, [JOB_ID]), [JOB_ID]);
  assert.deepEqual(await cancelAndScrubUserJobs(env, USER_ID, [JOB_ID]), [JOB_ID]);
  assert.equal(
    (await (await jobsBucket.get(`jobs/${JOB_ID}/cancelled.json`)).json()).ownerId,
    USER_ID
  );
});

test("failed job scrubbing retains its quota index until a retry succeeds", async () => {
  const jobsBucket = new FailFirstJobPayloadDeleteR2();
  const env = makeEnv(undefined, { JOBS_BUCKET: jobsBucket });
  const now = Date.now();
  await reserveJobSlot(env, USER_ID, JOB_ID, now);
  await jobsBucket.put(`jobs/${JOB_ID}/owner.json`, JSON.stringify({ ownerId: USER_ID }));
  await jobsBucket.put(`jobs/${JOB_ID}/status.json`, JSON.stringify({ status: "processing" }));
  await jobsBucket.put(`jobs/${JOB_ID}/input.pptx`, "sensitive-input");

  await assert.rejects(
    cleanupFailedJob(env, USER_ID, JOB_ID, [
      `jobs/${JOB_ID}/owner.json`,
      `jobs/${JOB_ID}/status.json`,
      `jobs/${JOB_ID}/input.pptx`,
    ]),
    /temporary job-bucket failure/
  );
  assert.deepEqual(await listUserJobIds(env, USER_ID), [JOB_ID]);

  await cleanupFailedJob(env, USER_ID, JOB_ID, []);
  assert.equal(await jobsBucket.get(`jobs/${JOB_ID}/input.pptx`), null);
  assert.deepEqual(await listUserJobIds(env, USER_ID), []);
});

test("failed account cleanup stays disabled and can be retried idempotently", async () => {
  const env = makeEnv(new FailFirstPlaybookDeleteR2());
  const salt = generateSaltHex();
  const account = await seedAccount(env, {
    extra: {
      salt,
      iterations: 1000,
      hash: await hashPassword("delete-password", salt, 1000),
    },
  });
  await env.PLAYBOOK_BUCKET.put(
    `accounts/${USER_ID}/playbook.json`,
    JSON.stringify({ schema: 1, offense: [], defense: [] })
  );
  const session = await sessionRequest(env, account);
  const makeDeleteRequest = (cookie = session.headers.get("cookie")) =>
    new Request("https://example.test/api/auth/delete-account", {
      method: "POST",
      headers: { cookie },
      body: JSON.stringify({ password: "delete-password", userId: USER_ID }),
    });

  const response = await deleteAccount({ request: makeDeleteRequest(), env });
  assert.equal(response.status, 503);
  assert.equal((await json(response)).deletionPending, true);
  assert.equal(await getUser(session, env), null);
  assert.equal((await getUser(session, env, { allowDisabled: true })).userId, USER_ID);
  assert.ok(await env.PLAYBOOK_BUCKET.get(`accounts/${USER_ID}/deletion.json`));

  const relogin = await login({
    request: new Request("https://example.test/api/auth/login", {
      method: "POST",
      body: JSON.stringify({ email: account.email, password: "delete-password" }),
    }),
    env,
  });
  assert.equal(relogin.status, 423);
  assert.equal((await relogin.clone().json()).deletionPending, true);
  assert.equal(relogin.headers.get("set-cookie"), null);

  const retried = await deleteAccount({
    request: new Request("https://example.test/api/auth/delete-account", {
      method: "POST",
      body: JSON.stringify({
        email: account.email,
        password: "delete-password",
        userId: USER_ID,
      }),
    }),
    env,
  });
  assert.equal(retried.status, 200);
  assert.equal(retried.headers.get("set-cookie"), null);
  const credentialTombstone = await env.PLAYBOOK_BUCKET.get(await emailKey(account.email));
  assert.ok((await credentialTombstone.json()).deletedAt);
  assert.equal(await env.PLAYBOOK_BUCKET.get(`accounts/${USER_ID}/deletion.json`), null);

  const replacement = await register({
    request: new Request("https://example.test/api/auth/register", {
      method: "POST",
      body: JSON.stringify({ email: account.email, password: "replacement-password" }),
    }),
    env,
  });
  assert.equal(replacement.status, 200);
  const replacementData = await replacement.json();
  assert.notEqual(replacementData.userId, USER_ID);

  const replacementCookie = replacement.headers.get("set-cookie").split(";", 1)[0];
  const staleDeletion = await deleteAccount({
    request: new Request("https://example.test/api/auth/delete-account", {
      method: "POST",
      headers: { cookie: replacementCookie },
      body: JSON.stringify({
        password: "replacement-password",
        userId: USER_ID,
      }),
    }),
    env,
  });
  assert.equal(staleDeletion.status, 409);
  assert.equal(
    (await (await env.PLAYBOOK_BUCKET.get(await emailKey(account.email))).json()).userId,
    replacementData.userId
  );
});

test("status hides jobs owned by another account", async () => {
  const env = makeEnv(undefined, { JOBS_BUCKET: new MemoryR2() });
  const account = await seedAccount(env);
  await env.JOBS_BUCKET.put(
    `jobs/${JOB_ID}/owner.json`,
    JSON.stringify({ ownerId: OTHER_USER_ID })
  );
  await env.JOBS_BUCKET.put(
    `jobs/${JOB_ID}/status.json`,
    JSON.stringify({ status: "processing" })
  );
  const request = await sessionRequest(env, account, `https://example.test/api/status/${JOB_ID}`);
  const response = await getStatus({ request, env, params: { jobId: JOB_ID } });

  assert.equal(response.status, 404);
  assert.equal(response.headers.get("cache-control"), "private, no-store");
});

test("status persistently expires stale processing jobs and releases their slot", async () => {
  const env = makeEnv(undefined, {
    JOBS_BUCKET: new MemoryR2(),
    MAX_ACTIVE_JOBS_PER_USER: "1",
    JOB_STALE_MINUTES: "10",
  });
  const account = await seedAccount(env);
  const now = Date.now();
  await reserveJobSlot(env, USER_ID, JOB_ID, now);
  await env.JOBS_BUCKET.put(
    `jobs/${JOB_ID}/owner.json`,
    JSON.stringify({ ownerId: USER_ID })
  );
  await env.JOBS_BUCKET.put(
    `jobs/${JOB_ID}/status.json`,
    JSON.stringify({ status: "processing", createdAt: new Date(now - 60 * 60 * 1000).toISOString() })
  );

  const request = await sessionRequest(env, account, `https://example.test/api/status/${JOB_ID}`);
  const response = await getStatus({ request, env, params: { jobId: JOB_ID } });
  assert.equal(response.status, 200);
  assert.equal((await json(response)).status, "error");
  assert.equal(
    (await (await env.JOBS_BUCKET.get(`jobs/${JOB_ID}/status.json`)).json()).status,
    "error"
  );
  const nextId = "00000000-0000-4000-8000-000000000099";
  assert.equal((await reserveJobSlot(env, USER_ID, nextId, now + 1000)).ok, true);
});

test("downloads require ownership, completion, and an explicitly listed basename", async () => {
  const env = makeEnv(undefined, { JOBS_BUCKET: new MemoryR2() });
  const account = await seedAccount(env);
  const filename = "offense_coach_card.pdf";
  await env.JOBS_BUCKET.put(
    `jobs/${JOB_ID}/owner.json`,
    JSON.stringify({ ownerId: USER_ID })
  );
  await env.JOBS_BUCKET.put(
    `jobs/${JOB_ID}/status.json`,
    JSON.stringify({ status: "complete", files: [filename] })
  );
  await env.JOBS_BUCKET.put(`jobs/${JOB_ID}/${filename}`, "%PDF-test");

  const request = await sessionRequest(
    env,
    account,
    `https://example.test/api/download/${JOB_ID}/${filename}`
  );
  const response = await download({
    request,
    env,
    params: { catchall: [JOB_ID, filename] },
  });
  assert.equal(response.status, 200);
  assert.equal(response.headers.get("cache-control"), "private, no-store");
  assert.equal(response.headers.get("x-content-type-options"), "nosniff");

  const missing = await download({
    request,
    env,
    params: { catchall: [JOB_ID, "unlisted.pdf"] },
  });
  assert.equal(missing.status, 404);
});

test("playbook saves return 409 when the conditional write loses a race", async () => {
  const env = makeEnv();
  const account = await seedAccount(env);
  const key = `accounts/${USER_ID}/playbook.json`;
  const base = "2026-07-11T12:00:00.000Z";
  await env.PLAYBOOK_BUCKET.put(
    key,
    JSON.stringify({ schema: 1, offense: [], defense: [], updatedAt: base })
  );
  env.PLAYBOOK_BUCKET.beforeConditionalPut = async (conditionalKey) => {
    if (conditionalKey !== key) return;
    await env.PLAYBOOK_BUCKET.put(
      key,
      JSON.stringify({ schema: 1, offense: [], defense: [], updatedAt: "newer" })
    );
  };

  const session = await sessionRequest(env, account);
  const request = new Request("https://example.test/api/plays", {
    method: "PUT",
    headers: { cookie: session.headers.get("cookie") },
    body: JSON.stringify({
      schema: 1,
      offense: [],
      defense: [],
      ownerId: USER_ID,
      baseUpdatedAt: base,
    }),
  });
  const response = await savePlays({ request, env });
  assert.equal(response.status, 409);
  assert.equal((await json(response)).serverUpdatedAt, "newer");
});

test("playbook saves reject a declared oversized body before parsing it", async () => {
  const env = makeEnv();
  const account = await seedAccount(env);
  const session = await sessionRequest(env, account);
  const request = new Request("https://example.test/api/plays", {
    method: "PUT",
    headers: {
      cookie: session.headers.get("cookie"),
      "content-length": "1000001",
    },
    body: "{}",
  });
  const response = await savePlays({ request, env });
  assert.equal(response.status, 413);
});

test("playbook saves reject a document captured for another account", async () => {
  const env = makeEnv();
  const account = await seedAccount(env);
  const session = await sessionRequest(env, account);
  const request = new Request("https://example.test/api/plays", {
    method: "PUT",
    headers: { cookie: session.headers.get("cookie") },
    body: JSON.stringify({
      schema: 1,
      offense: [],
      defense: [],
      ownerId: OTHER_USER_ID,
      baseUpdatedAt: null,
    }),
  });
  const response = await savePlays({ request, env });
  assert.equal(response.status, 403);
  assert.equal(await env.PLAYBOOK_BUCKET.get(`accounts/${USER_ID}/playbook.json`), null);
});

test("image generation rejects duplicate and zero-byte multipart files before dispatch", async () => {
  const env = makeEnv();
  const account = await seedAccount(env);
  const session = await sessionRequest(env, account);
  const pngHeader = new Uint8Array([0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a]);
  const options = JSON.stringify({ offense_coach_card: true });

  const duplicateForm = new FormData();
  duplicateForm.set("options", options);
  duplicateForm.append("plays", new File([pngHeader], "01.png", { type: "image/png" }));
  duplicateForm.append("plays", new File([pngHeader], "01.png", { type: "image/png" }));
  const duplicate = await generate({
    request: new Request("https://example.test/api/generate", {
      method: "POST",
      headers: { cookie: session.headers.get("cookie") },
      body: duplicateForm,
    }),
    env,
  });
  assert.equal(duplicate.status, 400);
  assert.match((await json(duplicate)).error, /Duplicate/);

  const emptyForm = new FormData();
  emptyForm.set("options", options);
  emptyForm.append("plays", new File([], "01.png", { type: "image/png" }));
  const empty = await generate({
    request: new Request("https://example.test/api/generate", {
      method: "POST",
      headers: { cookie: session.headers.get("cookie") },
      body: emptyForm,
    }),
    env,
  });
  assert.equal(empty.status, 400);
  assert.match((await json(empty)).error, /empty/);
});
