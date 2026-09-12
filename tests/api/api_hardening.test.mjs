import assert from "node:assert/strict";
import test from "node:test";

import {
  createRecoveryFields,
  createSessionCookie,
  emailKey,
  generateSaltHex,
  getUser,
  hashPassword,
  normalizeRecoveryCode,
  PASSWORD_ITERATIONS,
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
import { onRequestGet as getMe } from "../../dashboard/functions/api/auth/me.js";
import { onRequestPost as register } from "../../dashboard/functions/api/auth/register.js";
import {
  onRequestDelete as deletePlaybook,
  onRequestGet as getPlaybooks,
  onRequestPatch as renamePlaybook,
  onRequestPost as createPlaybook,
} from "../../dashboard/functions/api/playbooks.js";
import { onRequestPost as generate } from "../../dashboard/functions/api/generate.js";
import { onRequestPost as upload } from "../../dashboard/functions/api/upload.js";
import {
  onRequestGet as getPlays,
  onRequestPut as savePlays,
} from "../../dashboard/functions/api/plays.js";
import { onRequestGet as getStatus } from "../../dashboard/functions/api/status/[[jobId]].js";
import { onRequestGet as download } from "../../dashboard/functions/api/download/[[catchall]].js";
import {
  deleteAccountPlaybooks,
  playbookCatalogKey,
  playbookObjectKey,
} from "../../dashboard/functions/_lib/playbooks.js";

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
    this.beforeGetReturn = null;
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
      uploaded: new Date(),
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
      uploaded: stored.uploaded,
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
    if (this.beforeGetReturn) await this.beforeGetReturn(key);
    return stored ? this.object(key, stored) : null;
  }

  async head(key) {
    const stored = this.objects.get(key);
    return stored ? this.object(key, stored) : null;
  }

  async delete(keys) {
    for (const key of Array.isArray(keys) ? keys : [keys]) this.objects.delete(key);
  }

  async list({ prefix = "", limit = 1000, cursor } = {}) {
    const offset = cursor === undefined ? 0 : Number.parseInt(cursor, 10);
    const keys = [...this.objects.keys()]
      .filter((key) => key.startsWith(prefix))
      .sort();
    const objects = keys
      .slice(offset, offset + limit)
      .map((key) => ({ key }));
    const nextOffset = offset + objects.length;
    const truncated = nextOffset < keys.length;
    return {
      objects,
      truncated,
      ...(truncated ? { cursor: String(nextOffset) } : {}),
    };
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

class FailNamedPlaybookDeletesR2 extends MemoryR2 {
  constructor(failures = 1) {
    super();
    this.failuresRemaining = failures;
    this.namedDeleteAttempts = 0;
  }

  async delete(keys) {
    const list = Array.isArray(keys) ? keys : [keys];
    if (list.some((key) => key.includes("/playbooks/items/"))) {
      this.namedDeleteAttempts += 1;
    }
    if (
      this.failuresRemaining > 0 &&
      list.some((key) => key.includes("/playbooks/items/"))
    ) {
      this.failuresRemaining -= 1;
      throw new Error("temporary named-playbook delete failure");
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

class ThrowAfterCatalogPutR2 extends MemoryR2 {
  constructor() {
    super();
    this.throwAfterCatalogPut = false;
  }

  async put(key, value, options = {}) {
    const stored = await super.put(key, value, options);
    if (
      stored &&
      this.throwAfterCatalogPut &&
      key.endsWith("/playbooks/catalog.json")
    ) {
      this.throwAfterCatalogPut = false;
      throw new Error("ambiguous catalog transport failure");
    }
    return stored;
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

function requestWithResponseCookie(response, url = "https://example.test/api") {
  const setCookie = response.headers.get("set-cookie");
  assert.ok(setCookie, "expected an authenticated response cookie");
  return new Request(url, { headers: { cookie: setCookie.split(";", 1)[0] } });
}

// Accounts and cookies created before auth versioning had no sessionVersion,
// `sv`, or `iat`. Keep an exact old-token fixture so the fallback cannot be
// accidentally removed while initial pilot accounts still exist.
async function legacySessionRequest(env, account, url = "https://example.test/api") {
  const payload = Buffer.from(JSON.stringify({
    uid: account.userId,
    em: account.email,
    exp: Math.floor(Date.now() / 1000) + 3600,
  })).toString("base64url");
  const key = await crypto.subtle.importKey(
    "raw",
    Buffer.from(env.SESSION_SECRET, "hex"),
    { name: "HMAC", hash: "SHA-256" },
    false,
    ["sign"]
  );
  const signature = Buffer.from(
    await crypto.subtle.sign("HMAC", key, new TextEncoder().encode(payload))
  ).toString("base64url");
  return new Request(url, { headers: { cookie: `pb_session=${payload}.${signature}` } });
}

async function json(response) {
  return response.json();
}

function chips(keys) {
  return Object.fromEntries(keys.map((key, index) => [key, {
    x: (index + 1) / (keys.length + 1),
    y: 0.7,
  }]));
}

async function authenticatedRequest(env, account, url, init = {}) {
  const session = await sessionRequest(env, account);
  const headers = new Headers(init.headers || {});
  headers.set("cookie", session.headers.get("cookie"));
  return new Request(url, { ...init, headers });
}

async function callCreatePlaybook(env, account, name, defaultPlayersPerSide) {
  return createPlaybook({
    request: await authenticatedRequest(env, account, "https://example.test/api/playbooks", {
      method: "POST",
      body: JSON.stringify({ name, defaultPlayersPerSide }),
    }),
    env,
  });
}

async function callDeletePlaybook(env, account, playbookId, baseRevision) {
  return deletePlaybook({
    request: await authenticatedRequest(env, account, "https://example.test/api/playbooks", {
      method: "DELETE",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({ playbookId, baseRevision }),
    }),
    env,
  });
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

test("pre-version accounts and cookies remain valid and the account can log in", async () => {
  const env = makeEnv();
  const email = "initial.coach@example.com";
  const password = "original-password";
  const salt = generateSaltHex();
  const account = {
    userId: USER_ID,
    email,
    salt,
    iterations: 100000,
    hash: await hashPassword(password, salt, 100000),
    createdAt: "2026-07-07T19:47:02.000Z",
  };
  await env.PLAYBOOK_BUCKET.put(await emailKey(email), JSON.stringify(account));

  const legacyRequest = await legacySessionRequest(env, account);
  const legacyUser = await getUser(legacyRequest, env);
  assert.equal(legacyUser.userId, USER_ID);
  assert.equal(legacyUser.email, email);
  assert.equal(legacyUser.sessionVersion, 1);
  assert.equal(legacyUser.authenticatedAt, 0);
  const legacyMe = await getMe({ request: legacyRequest, env });
  assert.equal(legacyMe.status, 200);
  assert.deepEqual(await legacyMe.json(), { email, userId: USER_ID });

  const response = await login({
    request: new Request("https://example.test/api/auth/login", {
      method: "POST",
      body: JSON.stringify({ email, password }),
    }),
    env,
  });
  assert.equal(response.status, 200);
  assert.deepEqual(await response.clone().json(), { email, userId: USER_ID });
  assert.equal((await getUser(requestWithResponseCookie(response), env)).userId, USER_ID);

  const stored = await (await env.PLAYBOOK_BUCKET.get(await emailKey(email))).json();
  assert.deepEqual(stored, account, "a current-work-factor legacy login must not rewrite the account");
});

test("a successful legacy password login upgrades its PBKDF2 work factor", async () => {
  const env = makeEnv();
  const email = "low-work-factor@example.com";
  const password = "upgrade-password";
  const oldSalt = generateSaltHex();
  const account = {
    userId: USER_ID,
    email,
    salt: oldSalt,
    iterations: 1000,
    hash: await hashPassword(password, oldSalt, 1000),
    createdAt: "2026-01-01T00:00:00.000Z",
    compatibilityMarker: "preserve-me",
  };
  await env.PLAYBOOK_BUCKET.put(await emailKey(email), JSON.stringify(account));

  const response = await login({
    request: new Request("https://example.test/api/auth/login", {
      method: "POST",
      body: JSON.stringify({ email, password }),
    }),
    env,
  });
  assert.equal(response.status, 200);

  const upgraded = await (await env.PLAYBOOK_BUCKET.get(await emailKey(email))).json();
  assert.equal(upgraded.iterations, PASSWORD_ITERATIONS);
  assert.equal(upgraded.sessionVersion, 1);
  assert.notEqual(upgraded.salt, oldSalt);
  assert.equal(upgraded.compatibilityMarker, "preserve-me");
  assert.equal(
    upgraded.hash,
    await hashPassword(password, upgraded.salt, PASSWORD_ITERATIONS)
  );
  assert.equal((await getUser(requestWithResponseCookie(response), env)).userId, USER_ID);
});

test("new registration persists current credentials and returns a usable session", async () => {
  const env = makeEnv();
  const email = "new.coach@example.com";
  const password = "new-coach-password";
  const response = await register({
    request: new Request("https://example.test/api/auth/register", {
      method: "POST",
      headers: { "content-type": "application/json" },
      body: JSON.stringify({ email: `  ${email.toUpperCase()}  `, password }),
    }),
    env,
  });

  assert.equal(response.status, 200);
  const data = await response.clone().json();
  assert.equal(data.email, email);
  assert.match(data.userId, /^[0-9a-f-]{36}$/i);
  assert.match(data.recoveryCode, /^(?:[0-9A-F]{4}-){4}[0-9A-F]{4}$/);

  const stored = await (await env.PLAYBOOK_BUCKET.get(await emailKey(email))).json();
  assert.equal(stored.userId, data.userId);
  assert.equal(stored.email, email);
  assert.equal(stored.iterations, PASSWORD_ITERATIONS);
  assert.equal(stored.sessionVersion, 1);
  assert.equal(stored.hash, await hashPassword(password, stored.salt, PASSWORD_ITERATIONS));
  assert.equal(
    stored.recoveryHash,
    await hashPassword(
      normalizeRecoveryCode(data.recoveryCode),
      stored.recoverySalt,
      stored.recoveryIterations
    )
  );

  const registeredRequest = requestWithResponseCookie(
    response,
    "https://example.test/api/auth/me"
  );
  const registeredUser = await getUser(registeredRequest, env);
  assert.equal(registeredUser.userId, data.userId);
  assert.equal(registeredUser.email, email);
  const me = await getMe({ request: registeredRequest, env });
  assert.equal(me.status, 200);
  assert.deepEqual(await me.json(), { email, userId: data.userId });

  const relogin = await login({
    request: new Request("https://example.test/api/auth/login", {
      method: "POST",
      headers: { "content-type": "application/json" },
      body: JSON.stringify({ email, password }),
    }),
    env,
  });
  assert.equal(relogin.status, 200);
  assert.equal((await relogin.json()).userId, data.userId);
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
  assert.equal(updated.iterations, 100000);
});

test("password iterations stay within the Workers PBKDF2 limit", () => {
  // Workers WebCrypto throws NotSupportedError above 100000 iterations, which
  // breaks login/register/recover in production while Node-based tests pass.
  assert.ok(PASSWORD_ITERATIONS <= 100000);
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

test("status returns optional conversion warnings to the job owner", async () => {
  const env = makeEnv(undefined, { JOBS_BUCKET: new MemoryR2() });
  const account = await seedAccount(env);
  const warnings = [{
    code: "assumed_offense_before_defense",
    playCount: 2,
  }];
  await env.JOBS_BUCKET.put(
    `jobs/${JOB_ID}/owner.json`,
    JSON.stringify({ ownerId: USER_ID })
  );
  await env.JOBS_BUCKET.put(
    `jobs/${JOB_ID}/status.json`,
    JSON.stringify({
      status: "complete",
      files: ["offense_coach_card.pdf"],
      warnings,
    })
  );

  const request = await sessionRequest(env, account, `https://example.test/api/status/${JOB_ID}`);
  const response = await getStatus({ request, env, params: { jobId: JOB_ID } });

  assert.equal(response.status, 200);
  assert.deepEqual((await json(response)).warnings, warnings);
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

test("a new account receives an unconfigured schema-2 playbook", async () => {
  const env = makeEnv();
  const account = await seedAccount(env);
  const request = await sessionRequest(env, account, "https://example.test/api/plays");

  const response = await getPlays({ request, env });

  assert.equal(response.status, 200);
  assert.equal(response.headers.get("cache-control"), "private, no-store");
  assert.deepEqual(await json(response), {
    schema: 2,
    defaultPlayersPerSide: null,
    offense: [],
    defense: [],
  });
});

test("the playbook catalog adopts the legacy document without copying or duplicating it", async () => {
  const env = makeEnv();
  const account = await seedAccount(env);
  const legacyKey = `accounts/${USER_ID}/playbook.json`;
  const legacyDoc = {
    schema: 2,
    defaultPlayersPerSide: 6,
    offense: [],
    defense: [],
    updatedAt: "2026-08-01T12:00:00.000Z",
  };
  await env.PLAYBOOK_BUCKET.put(legacyKey, JSON.stringify(legacyDoc));

  const first = await getPlaybooks({
    request: await authenticatedRequest(env, account, "https://example.test/api/playbooks"),
    env,
  });
  assert.equal(first.status, 200);
  const firstBody = await json(first);
  assert.equal(firstBody.maxPlaybooks, 20);
  assert.equal(firstBody.playbooks.length, 1);
  assert.equal(firstBody.playbooks[0].id, "default");
  assert.equal(firstBody.playbooks[0].name, "My Playbook");
  assert.deepEqual(await (await env.PLAYBOOK_BUCKET.get(legacyKey)).json(), legacyDoc);

  const second = await getPlaybooks({
    request: await authenticatedRequest(env, account, "https://example.test/api/playbooks"),
    env,
  });
  const secondBody = await json(second);
  assert.deepEqual(secondBody.playbooks, firstBody.playbooks);
  assert.ok(await env.PLAYBOOK_BUCKET.get(playbookCatalogKey(USER_ID)));
});

test("an ambiguously committed catalog append is reconciled as a successful create", async () => {
  const bucket = new ThrowAfterCatalogPutR2();
  const env = makeEnv(bucket);
  const account = await seedAccount(env);
  const initialized = await getPlaybooks({
    request: await authenticatedRequest(env, account, "https://example.test/api/playbooks"),
    env,
  });
  assert.equal(initialized.status, 200);

  bucket.throwAfterCatalogPut = true;
  const created = await callCreatePlaybook(env, account, "Spring 2028", 5);
  assert.equal(created.status, 201);
  const body = await json(created);
  const matches = body.playbooks.filter((entry) => entry.name === "Spring 2028");
  assert.equal(matches.length, 1);
  assert.ok(await env.PLAYBOOK_BUCKET.get(playbookObjectKey(USER_ID, matches[0].id)));
  assert.equal((await env.PLAYBOOK_BUCKET.list({
    prefix: `accounts/${USER_ID}/playbooks/items/`,
  })).objects.length, 1);
});

test("catalog early returns remove data recreated during an account-deletion race", async () => {
  const bucket = new MemoryR2();
  const env = makeEnv(bucket);
  const account = await seedAccount(env);

  bucket.beforeConditionalPut = async (key) => {
    if (key !== playbookCatalogKey(USER_ID)) return;
    const credentialKey = await emailKey(account.email);
    const credential = await (await env.PLAYBOOK_BUCKET.get(credentialKey)).json();
    credential.disabledAt = new Date().toISOString();
    await env.PLAYBOOK_BUCKET.put(credentialKey, JSON.stringify(credential));
    await deleteAccountPlaybooks(env, USER_ID);
  };

  // The request authenticated before the hook disables the credential. Its
  // conditional create lands after the simulated deletion sweep, then the
  // duplicate-name early return must perform its own liveness cleanup.
  const raced = await callCreatePlaybook(env, account, "My Playbook", 5);
  assert.equal(raced.status, 409);
  assert.equal(await env.PLAYBOOK_BUCKET.get(playbookCatalogKey(USER_ID)), null);
  assert.equal((await env.PLAYBOOK_BUCKET.list({
    prefix: `accounts/${USER_ID}/playbooks/`,
  })).objects.length, 0);
});

test("an account can create, load, and save independent named 5v5 and 6v6 playbooks", async () => {
  const env = makeEnv();
  const account = await seedAccount(env);

  const [fiveResponse, sixResponse] = await Promise.all([
    callCreatePlaybook(env, account, "  Spring 2027 5v5  ", 5),
    callCreatePlaybook(env, account, "Fall 2027 6v6", 6),
  ]);
  assert.equal(fiveResponse.status, 201);
  assert.equal(sixResponse.status, 201);

  const catalogResponse = await getPlaybooks({
    request: await authenticatedRequest(env, account, "https://example.test/api/playbooks"),
    env,
  });
  const catalog = await json(catalogResponse);
  assert.equal(catalog.playbooks.length, 3);
  const five = catalog.playbooks.find((entry) => entry.name === "Spring 2027 5v5");
  const six = catalog.playbooks.find((entry) => entry.name === "Fall 2027 6v6");
  assert.ok(five);
  assert.ok(six);

  const load = async (id) => {
    const response = await getPlays({
      request: await authenticatedRequest(
        env,
        account,
        `https://example.test/api/plays?playbookId=${id}`
      ),
      env,
    });
    assert.equal(response.status, 200);
    return json(response);
  };
  const fiveDoc = await load(five.id);
  const sixDoc = await load(six.id);
  assert.equal(fiveDoc.defaultPlayersPerSide, 5);
  assert.equal(sixDoc.defaultPlayersPerSide, 6);

  const saveFive = await savePlays({
    request: await authenticatedRequest(
      env,
      account,
      `https://example.test/api/plays?playbookId=${five.id}`,
      {
        method: "PUT",
        body: JSON.stringify({
          ...fiveDoc,
          ownerId: USER_ID,
          baseRevision: fiveDoc.revision,
          offense: [{
            name: "Five only",
            playersPerSide: 5,
            chips: chips(["1", "2", "3", "C", "QB"]),
            routes: [],
          }],
        }),
      }
    ),
    env,
  });
  assert.equal(saveFive.status, 200);
  assert.equal((await load(five.id)).offense[0].name, "Five only");
  assert.deepEqual((await load(six.id)).offense, []);
  assert.equal(await env.PLAYBOOK_BUCKET.get(`accounts/${USER_ID}/playbook.json`), null);
});

test("named playbook saves reject stale revisions but separate books never conflict", async () => {
  const env = makeEnv();
  const account = await seedAccount(env);
  await callCreatePlaybook(env, account, "Book A", 5);
  await callCreatePlaybook(env, account, "Book B", 6);
  const listed = await getPlaybooks({
    request: await authenticatedRequest(env, account, "https://example.test/api/playbooks"),
    env,
  });
  const entries = (await json(listed)).playbooks;
  const a = entries.find((entry) => entry.name === "Book A");
  const b = entries.find((entry) => entry.name === "Book B");

  const getDocument = async (id) => json(await getPlays({
    request: await authenticatedRequest(
      env,
      account,
      `https://example.test/api/plays?playbookId=${id}`
    ),
    env,
  }));
  const originalA = await getDocument(a.id);
  const originalB = await getDocument(b.id);
  const save = async (id, document, name) => {
    const request = await authenticatedRequest(
      env,
      account,
      `https://example.test/api/plays?playbookId=${id}`,
      {
        method: "PUT",
        body: JSON.stringify({
          ...document,
          ownerId: USER_ID,
          baseRevision: document.revision,
          offense: [{
            name,
            playersPerSide: document.defaultPlayersPerSide,
            chips: document.defaultPlayersPerSide === 5
              ? chips(["1", "2", "3", "C", "QB"])
              : chips(["1", "2", "3", "4", "5", "QB"]),
            routes: [],
          }],
        }),
      }
    );
    return savePlays({ request, env });
  };

  const firstA = await save(a.id, originalA, "A wins");
  const firstB = await save(b.id, originalB, "B wins");
  assert.equal(firstA.status, 200);
  assert.equal(firstB.status, 200);

  const missingPrecondition = await savePlays({
    request: await authenticatedRequest(
      env,
      account,
      `https://example.test/api/plays?playbookId=${a.id}`,
      {
        method: "PUT",
        body: JSON.stringify({
          ...originalA,
          ownerId: USER_ID,
          offense: [{
            name: "No precondition",
            playersPerSide: 5,
            chips: chips(["1", "2", "3", "C", "QB"]),
            routes: [],
          }],
        }),
      }
    ),
    env,
  });
  assert.equal(missingPrecondition.status, 428);

  const staleA = await save(a.id, originalA, "A stale");
  assert.equal(staleA.status, 409);
  const conflict = await json(staleA);
  assert.equal(conflict.error, "conflict");
  assert.notEqual(conflict.serverRevision, originalA.revision);
  assert.equal((await getDocument(a.id)).offense[0].name, "A wins");
  assert.equal((await getDocument(b.id)).offense[0].name, "B wins");

  const forced = await savePlays({
    request: await authenticatedRequest(
      env,
      account,
      `https://example.test/api/plays?playbookId=${a.id}`,
      {
        method: "PUT",
        body: JSON.stringify({
          ...originalA,
          ownerId: USER_ID,
          forceOverwrite: true,
          offense: [{
            name: "Explicit overwrite",
            playersPerSide: 5,
            chips: chips(["1", "2", "3", "C", "QB"]),
            routes: [],
          }],
        }),
      }
    ),
    env,
  });
  assert.equal(forced.status, 200);
  assert.equal((await getDocument(a.id)).offense[0].name, "Explicit overwrite");
});

test("renaming a playbook is revision-safe and does not change its document", async () => {
  const env = makeEnv();
  const account = await seedAccount(env);
  const created = await callCreatePlaybook(env, account, "Winter 2027", 5);
  const entry = (await json(created)).playbooks.find((item) => item.name === "Winter 2027");
  const docBefore = await (await env.PLAYBOOK_BUCKET.get(
    playbookObjectKey(USER_ID, entry.id)
  )).json();

  const renamedResponse = await renamePlaybook({
    request: await authenticatedRequest(env, account, "https://example.test/api/playbooks", {
      method: "PATCH",
      body: JSON.stringify({
        playbookId: entry.id,
        name: "  Spring 2027  ",
        baseRevision: entry.revision,
      }),
    }),
    env,
  });
  assert.equal(renamedResponse.status, 200);
  const renamed = (await json(renamedResponse)).playbooks.find((item) => item.id === entry.id);
  assert.equal(renamed.name, "Spring 2027");
  assert.notEqual(renamed.revision, entry.revision);
  assert.deepEqual(
    await (await env.PLAYBOOK_BUCKET.get(playbookObjectKey(USER_ID, entry.id))).json(),
    docBefore
  );

  const staleRename = await renamePlaybook({
    request: await authenticatedRequest(env, account, "https://example.test/api/playbooks", {
      method: "PATCH",
      body: JSON.stringify({
        playbookId: entry.id,
        name: "Summer 2027",
        baseRevision: entry.revision,
      }),
    }),
    env,
  });
  assert.equal(staleRename.status, 409);
  assert.equal((await json(staleRename)).playbook.name, "Spring 2027");
});

test("deleting a named playbook removes only its catalog entry and document", async () => {
  const env = makeEnv();
  const account = await seedAccount(env);
  const firstResponse = await callCreatePlaybook(env, account, "Delete me", 5);
  const first = (await json(firstResponse)).playbooks.find((entry) => entry.name === "Delete me");
  const secondResponse = await callCreatePlaybook(env, account, "Keep me", 6);
  const second = (await json(secondResponse)).playbooks.find((entry) => entry.name === "Keep me");

  const deleted = await callDeletePlaybook(env, account, first.id, first.revision);
  assert.equal(deleted.status, 200);
  const catalog = await json(deleted);
  assert.equal(catalog.playbooks.some((entry) => entry.id === first.id), false);
  assert.equal(catalog.playbooks.some((entry) => entry.id === second.id), true);
  assert.equal(await env.PLAYBOOK_BUCKET.get(playbookObjectKey(USER_ID, first.id)), null);
  assert.ok(await env.PLAYBOOK_BUCKET.get(playbookObjectKey(USER_ID, second.id)));
  assert.ok(await env.PLAYBOOK_BUCKET.get(playbookObjectKey(USER_ID, "default")) === null);
});

test("playbook deletion rejects the default, malformed input, and a stale revision", async () => {
  const env = makeEnv();
  const account = await seedAccount(env);
  const initial = await getPlaybooks({
    request: await authenticatedRequest(env, account, "https://example.test/api/playbooks"),
    env,
  });
  const defaultEntry = (await json(initial)).playbooks[0];
  assert.equal(
    (await callDeletePlaybook(env, account, "default", defaultEntry.revision)).status,
    409
  );

  const malformed = await deletePlaybook({
    request: await authenticatedRequest(env, account, "https://example.test/api/playbooks", {
      method: "DELETE",
      body: "not-json",
    }),
    env,
  });
  assert.equal(malformed.status, 400);

  const created = await callCreatePlaybook(env, account, "Revision safe", 5);
  const entry = (await json(created)).playbooks.find((item) => item.name === "Revision safe");
  const renamed = await renamePlaybook({
    request: await authenticatedRequest(env, account, "https://example.test/api/playbooks", {
      method: "PATCH",
      body: JSON.stringify({
        playbookId: entry.id,
        name: "Renamed first",
        baseRevision: entry.revision,
      }),
    }),
    env,
  });
  assert.equal(renamed.status, 200);
  const stale = await callDeletePlaybook(env, account, entry.id, entry.revision);
  assert.equal(stale.status, 409);
  assert.equal((await json(stale)).playbook.name, "Renamed first");
  assert.ok(await env.PLAYBOOK_BUCKET.get(playbookObjectKey(USER_ID, entry.id)));
});

test("concurrent and repeated deletion is idempotent", async () => {
  const env = makeEnv();
  const account = await seedAccount(env);
  const created = await callCreatePlaybook(env, account, "Delete concurrently", 5);
  const entry = (await json(created)).playbooks.find(
    (item) => item.name === "Delete concurrently"
  );

  const responses = await Promise.all([
    callDeletePlaybook(env, account, entry.id, entry.revision),
    callDeletePlaybook(env, account, entry.id, entry.revision),
  ]);
  assert.deepEqual(responses.map((response) => response.status), [200, 200]);
  const repeated = await callDeletePlaybook(env, account, entry.id, entry.revision);
  assert.equal(repeated.status, 200);
  assert.equal(await env.PLAYBOOK_BUCKET.get(playbookObjectKey(USER_ID, entry.id)), null);
});

test("an ambiguously committed catalog removal completes playbook deletion", async () => {
  const bucket = new ThrowAfterCatalogPutR2();
  const env = makeEnv(bucket);
  const account = await seedAccount(env);
  const created = await callCreatePlaybook(env, account, "Ambiguous delete", 6);
  const entry = (await json(created)).playbooks.find((item) => item.name === "Ambiguous delete");

  bucket.throwAfterCatalogPut = true;
  const deleted = await callDeletePlaybook(env, account, entry.id, entry.revision);
  assert.equal(deleted.status, 200);
  assert.equal((await json(deleted)).playbooks.some((item) => item.id === entry.id), false);
  assert.equal(await env.PLAYBOOK_BUCKET.get(playbookObjectKey(USER_ID, entry.id)), null);
});

test("a transient document-delete failure is reconciled after catalog removal", async () => {
  const bucket = new FailNamedPlaybookDeletesR2();
  const env = makeEnv(bucket);
  const account = await seedAccount(env);
  const created = await callCreatePlaybook(env, account, "Retry cleanup", 5);
  const entry = (await json(created)).playbooks.find((item) => item.name === "Retry cleanup");

  const deleted = await callDeletePlaybook(env, account, entry.id, entry.revision);
  assert.equal(deleted.status, 200);
  assert.equal(bucket.failuresRemaining, 0);
  assert.equal(bucket.namedDeleteAttempts, 2);
  assert.equal((await json(deleted)).playbooks.some((item) => item.id === entry.id), false);
  assert.equal(await env.PLAYBOOK_BUCKET.get(playbookObjectKey(USER_ID, entry.id)), null);
});

test("repeated transient document-delete failures are retried after catalog removal", async () => {
  // Three failures exhaust the first bounded cleanup. The catch-path confirms
  // the catalog commit and starts another bounded cleanup, whose second try
  // succeeds. This covers the old two-failure orphan edge.
  const bucket = new FailNamedPlaybookDeletesR2(4);
  const env = makeEnv(bucket);
  const account = await seedAccount(env);
  const created = await callCreatePlaybook(env, account, "Persistent cleanup", 5);
  const entry = (await json(created)).playbooks.find((item) => item.name === "Persistent cleanup");

  const deleted = await callDeletePlaybook(env, account, entry.id, entry.revision);
  assert.equal(deleted.status, 200);
  assert.equal(bucket.failuresRemaining, 0);
  assert.equal(bucket.namedDeleteAttempts, 5);
  assert.equal((await json(deleted)).playbooks.some((item) => item.id === entry.id), false);
  assert.equal(await env.PLAYBOOK_BUCKET.get(playbookObjectKey(USER_ID, entry.id)), null);
});

test("deletion prevents an in-flight playbook save from resurrecting its document", async () => {
  const bucket = new MemoryR2();
  const env = makeEnv(bucket);
  const account = await seedAccount(env);
  const created = await callCreatePlaybook(env, account, "Delete during save", 5);
  const entry = (await json(created)).playbooks.find((item) => item.name === "Delete during save");
  const key = playbookObjectKey(USER_ID, entry.id);
  const original = await json(await getPlays({
    request: await authenticatedRequest(
      env,
      account,
      `https://example.test/api/plays?playbookId=${entry.id}`
    ),
    env,
  }));

  let releaseSave;
  let saveReadObject;
  const saveReadObjectPromise = new Promise((resolve) => { saveReadObject = resolve; });
  const releaseSavePromise = new Promise((resolve) => { releaseSave = resolve; });
  bucket.beforeGetReturn = async (readKey) => {
    if (readKey !== key) return;
    bucket.beforeGetReturn = null;
    saveReadObject();
    await releaseSavePromise;
  };
  const saving = savePlays({
    request: await authenticatedRequest(
      env,
      account,
      `https://example.test/api/plays?playbookId=${entry.id}`,
      {
        method: "PUT",
        body: JSON.stringify({
          ...original,
          ownerId: USER_ID,
          baseRevision: original.revision,
          offense: [{
            name: "Racing save",
            playersPerSide: 5,
            chips: chips(["1", "2", "3", "C", "QB"]),
            routes: [],
          }],
        }),
      }
    ),
    env,
  });
  await saveReadObjectPromise;
  const deleted = await callDeletePlaybook(env, account, entry.id, entry.revision);
  assert.equal(deleted.status, 200);
  releaseSave();
  assert.equal((await saving).status, 409);
  assert.equal(await env.PLAYBOOK_BUCKET.get(key), null);
});

test("playbook APIs reject invalid names, formats, malformed IDs, and unknown IDs", async () => {
  const env = makeEnv();
  const account = await seedAccount(env);
  for (const [name, count] of [["   ", 5], ["x".repeat(61), 5], ["Bad\nName", 5], ["Valid", 7]]) {
    const response = await callCreatePlaybook(env, account, name, count);
    assert.equal(response.status, 400);
  }

  const malformed = await getPlays({
    request: await authenticatedRequest(
      env,
      account,
      "https://example.test/api/plays?playbookId=../playbook.json"
    ),
    env,
  });
  assert.equal(malformed.status, 400);

  const unknownId = "99999999-9999-4999-8999-999999999999";
  const unknown = await getPlays({
    request: await authenticatedRequest(
      env,
      account,
      `https://example.test/api/plays?playbookId=${unknownId}`
    ),
    env,
  });
  assert.equal(unknown.status, 404);
  assert.equal(await env.PLAYBOOK_BUCKET.get(playbookObjectKey(USER_ID, unknownId)), null);

  const created = await callCreatePlaybook(env, account, "Spring", 5);
  assert.equal(created.status, 201);
  const duplicate = await callCreatePlaybook(env, account, "  spring  ", 6);
  assert.equal(duplicate.status, 409);

  const otherAccount = await seedAccount(env, {
    userId: OTHER_USER_ID,
    email: "other@example.com",
  });
  const otherCreated = await callCreatePlaybook(env, otherAccount, "Private", 6);
  const otherId = (await json(otherCreated)).playbooks.find(
    (entry) => entry.name === "Private"
  ).id;
  const foreign = await getPlays({
    request: await authenticatedRequest(
      env,
      account,
      `https://example.test/api/plays?playbookId=${otherId}`
    ),
    env,
  });
  assert.equal(foreign.status, 404);
});

test("the concurrent account cap keeps successful playbooks and creates no extra item", async () => {
  const env = makeEnv();
  const account = await seedAccount(env);
  const creations = [];
  for (let i = 1; i <= 19; i += 1) {
    creations.push(callCreatePlaybook(env, account, `Season ${i}`, i % 2 ? 5 : 6));
  }
  const results = await Promise.all(creations);
  assert.equal(results.filter((response) => response.status === 201).length, 19);

  const listed = await getPlaybooks({
    request: await authenticatedRequest(env, account, "https://example.test/api/playbooks"),
    env,
  });
  assert.equal((await json(listed)).playbooks.length, 20);
  const before = await env.PLAYBOOK_BUCKET.list({
    prefix: `accounts/${USER_ID}/playbooks/items/`,
  });
  assert.equal(before.objects.length, 19);

  const over = await callCreatePlaybook(env, account, "One too many", 5);
  assert.equal(over.status, 409);
  const after = await env.PLAYBOOK_BUCKET.list({
    prefix: `accounts/${USER_ID}/playbooks/items/`,
  });
  assert.equal(after.objects.length, 19);
});

test("account playbook cleanup paginates through cataloged and orphaned items only for that user", async () => {
  const env = makeEnv();
  await env.PLAYBOOK_BUCKET.put(`accounts/${USER_ID}/playbook.json`, "legacy");
  await Promise.all(Array.from({ length: 1005 }, (_, index) =>
    env.PLAYBOOK_BUCKET.put(
      `accounts/${USER_ID}/playbooks/items/orphan-${String(index).padStart(4, "0")}.json`,
      "sensitive"
    )
  ));
  await env.PLAYBOOK_BUCKET.put(
    `accounts/${OTHER_USER_ID}/playbooks/items/keep.json`,
    "other-user"
  );

  await deleteAccountPlaybooks(env, USER_ID);

  assert.equal(await env.PLAYBOOK_BUCKET.get(`accounts/${USER_ID}/playbook.json`), null);
  assert.equal((await env.PLAYBOOK_BUCKET.list({
    prefix: `accounts/${USER_ID}/playbooks/`,
  })).objects.length, 0);
  assert.ok(await env.PLAYBOOK_BUCKET.get(
    `accounts/${OTHER_USER_ID}/playbooks/items/keep.json`
  ));
});

test("legacy 5v5 defense N data migrates to 5 while 6v6 N and 5 stay unchanged", async () => {
  const env = makeEnv();
  const account = await seedAccount(env);
  const key = `accounts/${USER_ID}/playbook.json`;
  const legacyFiveChips = chips(["1", "2", "3", "4", "N"]);
  legacyFiveChips.N = { x: 0.47, y: 0.59 };
  const sixChips = chips(["1", "2", "3", "4", "5", "N"]);
  sixChips["5"] = { x: 0.79, y: 0.64 };
  sixChips.N = { x: 0.51, y: 0.61 };
  const legacyDoc = {
    schema: 2,
    defaultPlayersPerSide: 5,
    offense: [],
    defense: [{
      name: "Legacy five",
      playersPerSide: 5,
      chips: legacyFiveChips,
      routes: [{ chip: "N", points: [[0.47, 0.59], [0.47, 0.25]] }],
    }, {
      name: "Six stays six",
      playersPerSide: 6,
      chips: sixChips,
      routes: [
        { chip: "N", points: [[0.51, 0.61], [0.50, 0.30]] },
        { chip: "5", points: [[0.79, 0.64], [0.80, 0.35]] },
      ],
    }],
    updatedAt: "2026-09-02T12:00:00.000Z",
  };
  await env.PLAYBOOK_BUCKET.put(key, JSON.stringify(legacyDoc));

  const getRequest = await sessionRequest(env, account, "https://example.test/api/plays");
  const getResponse = await getPlays({ request: getRequest, env });
  const loaded = await json(getResponse);

  assert.equal(getResponse.status, 200);
  assert.deepEqual(Object.keys(loaded.defense[0].chips).sort(), ["1", "2", "3", "4", "5"]);
  assert.deepEqual(loaded.defense[0].chips["5"], { x: 0.47, y: 0.59 });
  assert.equal(Object.hasOwn(loaded.defense[0].chips, "N"), false);
  assert.equal(loaded.defense[0].routes[0].chip, "5");
  assert.deepEqual(loaded.defense[0].routes[0].points, [[0.47, 0.59], [0.47, 0.25]]);
  assert.deepEqual(loaded.defense[1].chips["5"], { x: 0.79, y: 0.64 });
  assert.deepEqual(loaded.defense[1].chips.N, { x: 0.51, y: 0.61 });
  assert.deepEqual(loaded.defense[1].routes.map((route) => route.chip), ["N", "5"]);

  // A browser tab loaded before this rollout can still submit the old shape;
  // the API canonicalizes it before validation and stores only the new key.
  const session = await sessionRequest(env, account);
  const putResponse = await savePlays({
    request: new Request("https://example.test/api/plays", {
      method: "PUT",
      headers: { cookie: session.headers.get("cookie") },
      body: JSON.stringify({ ...legacyDoc, ownerId: USER_ID }),
    }),
    env,
  });
  assert.equal(putResponse.status, 200);

  const stored = await (await env.PLAYBOOK_BUCKET.get(key)).json();
  assert.deepEqual(Object.keys(stored.defense[0].chips).sort(), ["1", "2", "3", "4", "5"]);
  assert.deepEqual(stored.defense[0].chips["5"], { x: 0.47, y: 0.59 });
  assert.equal(stored.defense[0].routes[0].chip, "5");
  assert.deepEqual(stored.defense[1].chips["5"], { x: 0.79, y: 0.64 });
  assert.deepEqual(stored.defense[1].chips.N, { x: 0.51, y: 0.61 });
  assert.deepEqual(stored.defense[1].routes.map((route) => route.chip), ["N", "5"]);
});

test("playbook saves preserve schema-1 compatibility and accept mixed schema-2 formats", async () => {
  const env = makeEnv();
  const account = await seedAccount(env);
  const session = await sessionRequest(env, account);
  const cookie = session.headers.get("cookie");
  const save = (body) => savePlays({
    request: new Request("https://example.test/api/plays", {
      method: "PUT",
      headers: { cookie },
      body: JSON.stringify(body),
    }),
    env,
  });

  const legacy = await save({
    schema: 1,
    offense: [{ name: "Legacy six", chips: chips(["1", "2", "3", "4", "5", "QB"]) }],
    defense: [],
    ownerId: USER_ID,
  });
  assert.equal(legacy.status, 200);

  const mixed = await save({
    schema: 2,
    defaultPlayersPerSide: 5,
    offense: [{
      name: "Five-player offense",
      playersPerSide: 5,
      chips: chips(["1", "2", "3", "C", "QB"]),
      routes: [{ chip: "C", points: [[0.5, 0.7], [0.5, 0.3]] }],
    }],
    defense: [{
      name: "Archived six-player defense",
      playersPerSide: 6,
      chips: chips(["1", "2", "3", "4", "5", "N"]),
      routes: [],
    }],
    ownerId: USER_ID,
  });
  assert.equal(mixed.status, 200);

  const stored = await (
    await env.PLAYBOOK_BUCKET.get(`accounts/${USER_ID}/playbook.json`)
  ).json();
  assert.equal(stored.schema, 2);
  assert.equal(stored.defaultPlayersPerSide, 5);
  assert.equal(stored.offense[0].playersPerSide, 5);
  assert.equal(stored.defense[0].playersPerSide, 6);
  assert.deepEqual(Object.keys(stored.offense[0].chips), ["1", "2", "3", "C", "QB"]);
});

test("schema-2 playbook saves reject invalid player formats and orphaned routes", async () => {
  const env = makeEnv();
  const account = await seedAccount(env);
  const session = await sessionRequest(env, account);
  const save = (play, defaultPlayersPerSide = 5) => savePlays({
    request: new Request("https://example.test/api/plays", {
      method: "PUT",
      headers: { cookie: session.headers.get("cookie") },
      body: JSON.stringify({
        schema: 2,
        defaultPlayersPerSide,
        offense: [play],
        defense: [],
        ownerId: USER_ID,
      }),
    }),
    env,
  });

  const wrongLineup = await save({
    name: "Wrong lineup",
    playersPerSide: 5,
    chips: chips(["1", "2", "3", "4", "QB"]),
    routes: [],
  });
  assert.equal(wrongLineup.status, 400);

  const orphanedRoute = await save({
    name: "Orphaned route",
    playersPerSide: 5,
    chips: chips(["1", "2", "3", "C", "QB"]),
    routes: [{ chip: "4", points: [[0.5, 0.7], [0.5, 0.3]] }],
  });
  assert.equal(orphanedRoute.status, 400);

  const invalidDefault = await save({
    name: "Invalid default",
    playersPerSide: 5,
    chips: chips(["1", "2", "3", "C", "QB"]),
    routes: [],
  }, 7);
  assert.equal(invalidDefault.status, 400);
});

test("a stale schema-1 editor cannot downgrade a schema-2 playbook", async () => {
  const env = makeEnv();
  const account = await seedAccount(env);
  const session = await sessionRequest(env, account);
  const cookie = session.headers.get("cookie");
  const save = (body) => savePlays({
    request: new Request("https://example.test/api/plays", {
      method: "PUT",
      headers: { cookie },
      body: JSON.stringify({ ...body, ownerId: USER_ID }),
    }),
    env,
  });

  const current = await save({
    schema: 2,
    defaultPlayersPerSide: 5,
    offense: [{
      name: "Center choice",
      playersPerSide: 5,
      chips: chips(["1", "2", "3", "C", "QB"]),
      routes: [{ chip: "C", points: [[0.5, 0.7], [0.5, 0.3]] }],
    }],
    defense: [],
  });
  assert.equal(current.status, 200);

  // Omit baseUpdatedAt to model the old editor's explicit "Overwrite server"
  // path. Schema protection must hold even when concurrency checks are waived.
  const stale = await save({
    schema: 1,
    offense: [{
      name: "Old six-player copy",
      chips: chips(["1", "2", "3", "4", "5", "QB"]),
      routes: [],
    }],
    defense: [],
  });
  assert.equal(stale.status, 422);
  assert.match((await json(stale)).error, /out of date/i);

  const stored = await (
    await env.PLAYBOOK_BUCKET.get(`accounts/${USER_ID}/playbook.json`)
  ).json();
  assert.equal(stored.schema, 2);
  assert.equal(stored.offense[0].playersPerSide, 5);
  assert.ok(Object.hasOwn(stored.offense[0].chips, "C"));
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

test("PPTX upload rejects signed-out requests before parsing or writing the file", async () => {
  const env = makeEnv();
  let parsed = false;
  const request = {
    headers: new Headers({ "content-type": "multipart/form-data; boundary=test" }),
    async formData() {
      parsed = true;
      throw new Error("signed-out uploads must never be parsed");
    },
  };

  const response = await upload({ request, env });
  assert.equal(response.status, 401);
  assert.deepEqual(await json(response), { error: "Not signed in" });
  assert.equal(parsed, false);
  assert.equal(env.PLAYBOOK_BUCKET.objects.size, 0);
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
