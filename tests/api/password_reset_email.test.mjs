import assert from "node:assert/strict";
import test from "node:test";

import {
  createRecoveryFields,
  createSessionCookie,
  emailKey,
  generateSaltHex,
  getUser,
  hashPassword,
  hashPasswordResetToken,
  normalizeRecoveryCode,
  PASSWORD_ITERATIONS,
} from "../../dashboard/functions/_lib/auth.js";
import {
  PASSWORD_RESET_GENERIC_MESSAGE,
  PASSWORD_RESET_TTL_MS,
} from "../../dashboard/functions/_lib/password-reset.js";
import { onRequestPost as requestReset } from
  "../../dashboard/functions/api/auth/password-reset/request.js";
import { onRequestPost as completeReset } from
  "../../dashboard/functions/api/auth/password-reset/complete.js";
import { onRequestPost as recoverWithCode } from
  "../../dashboard/functions/api/auth/recover.js";
import { onRequestPost as rotateRecoveryCode } from
  "../../dashboard/functions/api/auth/recovery-code.js";
import { onRequestPost as login } from
  "../../dashboard/functions/api/auth/login.js";

const SESSION_SECRET = "0123456789abcdef".repeat(4);
const USER_ID = "11111111-1111-4111-8111-111111111111";
const EMAIL = "coach@example.com";

async function bytes(value) {
  if (typeof value === "string") return new TextEncoder().encode(value);
  if (value instanceof ArrayBuffer) return new Uint8Array(value);
  if (ArrayBuffer.isView(value)) {
    return new Uint8Array(value.buffer.slice(value.byteOffset, value.byteOffset + value.byteLength));
  }
  throw new TypeError("Unsupported test value");
}

class MemoryR2 {
  constructor() {
    this.objects = new Map();
    this.sequence = 0;
    this.getCount = 0;
    this.blockGet = null;
    this.throwBeforePasswordResetPut = false;
    this.throwAfterPasswordResetPut = false;
    this.throwAfterCompletionPut = false;
  }

  object(key, stored) {
    const value = stored.value.slice();
    return {
      key,
      etag: stored.etag,
      async text() { return new TextDecoder().decode(value); },
      async json() { return JSON.parse(new TextDecoder().decode(value)); },
    };
  }

  async get(key) {
    this.getCount += 1;
    if (this.blockGet) await this.blockGet;
    const stored = this.objects.get(key);
    return stored ? this.object(key, stored) : null;
  }

  async put(key, value, options = {}) {
    const current = this.objects.get(key);
    const condition = options.onlyIf;
    if (condition) {
      if (condition.etagDoesNotMatch === "*" && current) return null;
      if (condition.etagMatches !== undefined && (!current || current.etag !== condition.etagMatches)) {
        return null;
      }
    }
    const valueBytes = await bytes(value);
    const parsed = JSON.parse(new TextDecoder().decode(valueBytes));
    if (this.throwBeforePasswordResetPut && parsed.passwordReset) {
      this.throwBeforePasswordResetPut = false;
      throw new Error("transient issue storage failure");
    }
    const stored = { value: valueBytes, etag: `etag-${++this.sequence}` };
    this.objects.set(key, stored);
    if (this.throwAfterPasswordResetPut && parsed.passwordReset) {
      this.throwAfterPasswordResetPut = false;
      throw new Error("ambiguous issue transport failure");
    }
    if (this.throwAfterCompletionPut && parsed.lastPasswordResetId && !parsed.passwordReset) {
      this.throwAfterCompletionPut = false;
      throw new Error("ambiguous completion transport failure");
    }
    return this.object(key, stored);
  }
}

class FakeEmailService {
  constructor() {
    this.allowed = true;
    this.permitCalls = [];
    this.messages = [];
    this.sendStatus = 204;
  }

  async fetch(input, init = {}) {
    const path = new URL(typeof input === "string" ? input : input.url).pathname;
    const body = JSON.parse(init.body || "{}");
    if (path === "/password-reset/permit") {
      this.permitCalls.push(body);
      return Response.json({ allowed: this.allowed });
    }
    if (path === "/password-reset") {
      this.messages.push(body);
      return new Response(null, { status: this.sendStatus });
    }
    return new Response(null, { status: 404 });
  }
}

function makeEnv(bucket = new MemoryR2(), emailService = new FakeEmailService()) {
  return { PLAYBOOK_BUCKET: bucket, EMAIL_SERVICE: emailService, SESSION_SECRET };
}

async function seedAccount(env, password = "old-password") {
  const salt = generateSaltHex();
  const { recoveryCode, fields } = await createRecoveryFields();
  const record = {
    userId: USER_ID,
    email: EMAIL,
    salt,
    iterations: PASSWORD_ITERATIONS,
    hash: await hashPassword(password, salt, PASSWORD_ITERATIONS),
    ...fields,
    sessionVersion: 1,
    createdAt: new Date().toISOString(),
  };
  await env.PLAYBOOK_BUCKET.put(await emailKey(EMAIL), JSON.stringify(record));
  return { record, recoveryCode };
}

function resetRequest(email = EMAIL, headers = {}) {
  return new Request("https://example.test/api/auth/password-reset/request", {
    method: "POST",
    headers: { "content-type": "application/json", "cf-connecting-ip": "192.0.2.1", ...headers },
    body: JSON.stringify({ email }),
  });
}

async function callResetRequest(env, email = EMAIL, { drain = true } = {}) {
  const work = [];
  const response = await requestReset({
    request: resetRequest(email),
    env,
    waitUntil(promise) { work.push(promise); },
  });
  if (drain) await Promise.all(work);
  return { response, work };
}

function completionRequest(email, token, newPassword = "new-password", headers = {}) {
  return new Request("https://example.test/api/auth/password-reset/complete", {
    method: "POST",
    headers: { "content-type": "application/json", ...headers },
    body: JSON.stringify({ email, token, newPassword }),
  });
}

function chunkedRequest(url, chunks, headers = {}) {
  const encoder = new TextEncoder();
  const body = new ReadableStream({
    start(controller) {
      for (const chunk of chunks) controller.enqueue(encoder.encode(chunk));
      controller.close();
    },
  });
  const request = new Request(url, { method: "POST", headers, body, duplex: "half" });
  assert.equal(request.headers.get("content-length"), null);
  return request;
}

async function currentRecord(env) {
  return (await env.PLAYBOOK_BUCKET.get(await emailKey(EMAIL))).json();
}

test("reset requests are uniform and keep account/provider work in waitUntil", async () => {
  const knownEnv = makeEnv();
  await seedAccount(knownEnv);
  let releaseGet;
  knownEnv.PLAYBOOK_BUCKET.blockGet = new Promise((resolve) => { releaseGet = resolve; });

  const known = await callResetRequest(knownEnv, EMAIL, { drain: false });
  assert.equal(known.response.status, 202);
  assert.equal(known.work.length, 1);
  assert.deepEqual(await known.response.clone().json(), { message: PASSWORD_RESET_GENERIC_MESSAGE });
  assert.match(known.response.headers.get("cache-control"), /no-store/);

  const missingEnv = makeEnv();
  const missing = await callResetRequest(missingEnv, "missing@example.com");
  assert.equal(missing.response.status, known.response.status);
  assert.equal(await missing.response.text(), await known.response.text());
  assert.equal(missing.response.headers.get("set-cookie"), null);

  releaseGet();
  await Promise.all(known.work);
});

test("a reset request stores only a digest and emails the committed 15-minute token", async () => {
  const env = makeEnv();
  const { record } = await seedAccount(env);
  const before = Date.now();
  const { response } = await callResetRequest(env);
  const after = Date.now();

  assert.equal(response.status, 202);
  assert.equal(env.EMAIL_SERVICE.messages.length, 1);
  const message = env.EMAIL_SERVICE.messages[0];
  assert.deepEqual(Object.keys(message).sort(), ["email", "token"]);
  assert.equal(message.email, EMAIL);
  assert.match(message.token, /^[A-Za-z0-9_-]{43}$/);

  const stored = await currentRecord(env);
  assert.equal(stored.passwordReset.tokenHash, await hashPasswordResetToken(message.token));
  assert.equal(stored.passwordReset.userId, USER_ID);
  assert.equal(stored.passwordReset.issuedForSessionVersion, record.sessionVersion);
  assert.equal(stored.sessionVersion, record.sessionVersion);
  assert.equal(stored.hash, record.hash);
  assert.equal(JSON.stringify(stored).includes(message.token), false);
  const expiry = Date.parse(stored.passwordReset.expiresAt);
  assert.ok(expiry >= before + PASSWORD_RESET_TTL_MS);
  assert.ok(expiry <= after + PASSWORD_RESET_TTL_MS);
});

test("permit, account cooldown, and hourly quota suppress delivery without changing the 202", async () => {
  const env = makeEnv();
  await seedAccount(env);

  env.EMAIL_SERVICE.allowed = false;
  const denied = await callResetRequest(env);
  assert.equal(denied.response.status, 202);
  assert.equal(env.PLAYBOOK_BUCKET.getCount, 0);
  assert.equal(env.EMAIL_SERVICE.messages.length, 0);

  env.EMAIL_SERVICE.allowed = true;
  await callResetRequest(env);
  assert.equal(env.EMAIL_SERVICE.messages.length, 1);
  await callResetRequest(env);
  assert.equal(env.EMAIL_SERVICE.messages.length, 1);

  const key = await emailKey(EMAIL);
  const object = await env.PLAYBOOK_BUCKET.get(key);
  const record = await object.json();
  record.passwordResetRate = {
    schema: 1,
    windowStartedAt: new Date(Date.now() - 10 * 60 * 1000).toISOString(),
    sendCount: 3,
    lastSentAt: new Date(Date.now() - 2 * 60 * 1000).toISOString(),
  };
  await env.PLAYBOOK_BUCKET.put(key, JSON.stringify(record));
  await callResetRequest(env);
  assert.equal(env.EMAIL_SERVICE.messages.length, 1);

  const dailyEnv = makeEnv();
  await seedAccount(dailyEnv);
  const dailyKey = await emailKey(EMAIL);
  const dailyObject = await dailyEnv.PLAYBOOK_BUCKET.get(dailyKey);
  const dailyRecord = await dailyObject.json();
  dailyRecord.passwordResetRate = {
    schema: 1,
    windowStartedAt: new Date(Date.now() - 2 * 60 * 60 * 1000).toISOString(),
    sendCount: 0,
    lastSentAt: new Date(Date.now() - 2 * 60 * 1000).toISOString(),
    dayStartedAt: new Date(Date.now() - 12 * 60 * 60 * 1000).toISOString(),
    daySendCount: 5,
  };
  await dailyEnv.PLAYBOOK_BUCKET.put(dailyKey, JSON.stringify(dailyRecord));
  const daily = await callResetRequest(dailyEnv);
  assert.equal(daily.response.status, 202);
  assert.equal(dailyEnv.EMAIL_SERVICE.messages.length, 0);
});

test("missing email binding fails before account lookup", async () => {
  const bucket = new MemoryR2();
  const env = { PLAYBOOK_BUCKET: bucket, SESSION_SECRET };
  const response = await requestReset({ request: resetRequest(), env, waitUntil() {} });
  assert.equal(response.status, 503);
  assert.equal(bucket.getCount, 0);
});

test("ambiguous challenge issuance is recognized and still sends exactly one matching link", async () => {
  const bucket = new MemoryR2();
  const env = makeEnv(bucket);
  await seedAccount(env);
  bucket.throwAfterPasswordResetPut = true;
  await callResetRequest(env);

  assert.equal(env.EMAIL_SERVICE.messages.length, 1);
  const stored = await currentRecord(env);
  assert.equal(
    stored.passwordReset.tokenHash,
    await hashPasswordResetToken(env.EMAIL_SERVICE.messages[0].token)
  );
});

test("a transient challenge write failure retries and emails only the committed token", async () => {
  const bucket = new MemoryR2();
  const env = makeEnv(bucket);
  await seedAccount(env);
  bucket.throwBeforePasswordResetPut = true;

  const { response } = await callResetRequest(env);
  assert.equal(response.status, 202);
  assert.equal(env.EMAIL_SERVICE.messages.length, 1);

  const stored = await currentRecord(env);
  assert.equal(
    stored.passwordReset.tokenHash,
    await hashPasswordResetToken(env.EMAIL_SERVICE.messages[0].token)
  );
});

test("valid email reset rotates credentials and revokes sessions without auto-login", async () => {
  const env = makeEnv();
  const { record } = await seedAccount(env);
  const oldCookie = await createSessionCookie(USER_ID, EMAIL, env, record.sessionVersion);
  const oldRequest = new Request("https://example.test/api", {
    headers: { cookie: oldCookie.split(";", 1)[0] },
  });
  await callResetRequest(env);
  const token = env.EMAIL_SERVICE.messages[0].token;

  const response = await completeReset({
    request: completionRequest(EMAIL, token, "new-password", {
      cookie: "pb_session=different-account-session",
    }),
    env,
  });
  assert.equal(response.status, 200);
  assert.match(response.headers.get("set-cookie"), /^pb_session=;/);
  assert.match(response.headers.get("set-cookie"), /Max-Age=0/);
  const data = await response.json();
  assert.match(data.recoveryCode, /^(?:[0-9A-F]{4}-){4}[0-9A-F]{4}$/);

  const stored = await currentRecord(env);
  assert.equal(stored.sessionVersion, 2);
  assert.equal(stored.passwordReset, undefined);
  assert.equal(stored.hash, await hashPassword("new-password", stored.salt, stored.iterations));
  assert.equal(
    stored.recoveryHash,
    await hashPassword(
      normalizeRecoveryCode(data.recoveryCode),
      stored.recoverySalt,
      stored.recoveryIterations
    )
  );
  assert.equal(await getUser(oldRequest, env), null);

  const relogin = await login({
    request: new Request("https://example.test/api/auth/login", {
      method: "POST",
      body: JSON.stringify({ email: EMAIL, password: "new-password" }),
    }),
    env,
  });
  assert.equal(relogin.status, 200);
});

test("wrong, expired, replayed, and concurrent reset links cannot mutate twice", async () => {
  const env = makeEnv();
  await seedAccount(env);
  await callResetRequest(env);
  const token = env.EMAIL_SERVICE.messages[0].token;

  const wrong = await completeReset({
    request: completionRequest(EMAIL, `${token.slice(0, -1)}${token.endsWith("A") ? "B" : "A"}`),
    env,
  });
  assert.equal(wrong.status, 401);

  const key = await emailKey(EMAIL);
  const object = await env.PLAYBOOK_BUCKET.get(key);
  const expiring = await object.json();
  expiring.passwordReset.expiresAt = new Date(Date.now() - 1).toISOString();
  await env.PLAYBOOK_BUCKET.put(key, JSON.stringify(expiring));
  const expired = await completeReset({ request: completionRequest(EMAIL, token), env });
  assert.equal(expired.status, 401);

  // Issue a fresh token outside the account cooldown.
  const stale = await currentRecord(env);
  stale.passwordResetRate.lastSentAt = new Date(Date.now() - 2 * 60 * 1000).toISOString();
  await env.PLAYBOOK_BUCKET.put(key, JSON.stringify(stale));
  await callResetRequest(env);
  const fresh = env.EMAIL_SERVICE.messages.at(-1).token;
  const [one, two] = await Promise.all([
    completeReset({ request: completionRequest(EMAIL, fresh, "winner-one"), env }),
    completeReset({ request: completionRequest(EMAIL, fresh, "winner-two"), env }),
  ]);
  assert.deepEqual([one.status, two.status].sort((a, b) => a - b), [200, 409]);
  const replay = await completeReset({ request: completionRequest(EMAIL, fresh, "third-password"), env });
  assert.equal(replay.status, 401);
  assert.equal((await currentRecord(env)).sessionVersion, 2);
});

test("manual recovery and recovery-code rotation revoke pending email links", async () => {
  const env = makeEnv();
  const { record, recoveryCode } = await seedAccount(env);
  await callResetRequest(env);
  const firstToken = env.EMAIL_SERVICE.messages.at(-1).token;
  const recovered = await recoverWithCode({
    request: new Request("https://example.test/api/auth/recover", {
      method: "POST",
      body: JSON.stringify({ email: EMAIL, recoveryCode, newPassword: "manual-reset" }),
    }),
    env,
  });
  assert.equal(recovered.status, 200);
  assert.equal(
    (await completeReset({ request: completionRequest(EMAIL, firstToken), env })).status,
    401
  );

  const key = await emailKey(EMAIL);
  const afterManual = await currentRecord(env);
  afterManual.passwordResetRate.lastSentAt = new Date(Date.now() - 2 * 60 * 1000).toISOString();
  await env.PLAYBOOK_BUCKET.put(key, JSON.stringify(afterManual));
  await callResetRequest(env);
  const secondToken = env.EMAIL_SERVICE.messages.at(-1).token;
  const sessionCookie = await createSessionCookie(
    USER_ID,
    EMAIL,
    env,
    (await currentRecord(env)).sessionVersion
  );
  const rotated = await rotateRecoveryCode({
    request: new Request("https://example.test/api/auth/recovery-code", {
      method: "POST",
      headers: { cookie: sessionCookie.split(";", 1)[0] },
    }),
    env,
  });
  assert.equal(rotated.status, 200);
  assert.equal(
    (await completeReset({ request: completionRequest(EMAIL, secondToken), env })).status,
    401
  );
  assert.equal((await currentRecord(env)).passwordReset, undefined);
  assert.equal(record.userId, USER_ID);
});

test("an ambiguously committed password reset is reported as success once", async () => {
  const bucket = new MemoryR2();
  const env = makeEnv(bucket);
  await seedAccount(env);
  await callResetRequest(env);
  const token = env.EMAIL_SERVICE.messages[0].token;
  bucket.throwAfterCompletionPut = true;

  const response = await completeReset({
    request: completionRequest(EMAIL, token, "ambiguous-password"),
    env,
  });
  assert.equal(response.status, 200);
  assert.match((await response.json()).recoveryCode, /^(?:[0-9A-F]{4}-){4}[0-9A-F]{4}$/);
  assert.equal((await currentRecord(env)).sessionVersion, 2);
});

test("password reset request and completion bound their JSON bodies", async () => {
  const env = makeEnv();
  const oversizedRequest = await requestReset({
    request: new Request("https://example.test/api/auth/password-reset/request", {
      method: "POST",
      headers: { "content-length": "2049" },
      body: "{}",
    }),
    env,
    waitUntil() {},
  });
  assert.equal(oversizedRequest.status, 413);

  const oversizedChunkedRequest = await requestReset({
    request: chunkedRequest(
      "https://example.test/api/auth/password-reset/request",
      ["{\"email\":\"", "x".repeat(2048), "\"}"]
    ),
    env,
    waitUntil() {},
  });
  assert.equal(oversizedChunkedRequest.status, 413);

  const oversizedChunkedCompletion = await completeReset({
    request: chunkedRequest(
      "https://example.test/api/auth/password-reset/complete",
      ["{\"newPassword\":\"", "x".repeat(10000), "\"}"]
    ),
    env,
  });
  assert.equal(oversizedChunkedCompletion.status, 413);

  const malformed = await completeReset({
    request: new Request("https://example.test/api/auth/password-reset/complete", {
      method: "POST",
      body: "not-json",
    }),
    env,
  });
  assert.equal(malformed.status, 400);
});
