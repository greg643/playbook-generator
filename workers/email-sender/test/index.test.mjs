import assert from "node:assert/strict";
import test from "node:test";

import { handleRequest } from "../src/index.js";

const RESET_TOKEN = "A".repeat(43);
const RATE_LIMIT_KEY = "a".repeat(64);

function request(path, body, options = {}) {
  const headers = new Headers(options.headers);
  if (!options.omitContentType) headers.set("content-type", "application/json");
  return new Request(`https://email.internal${path}`, {
    method: options.method || "POST",
    headers,
    body: options.method === "GET" ? undefined : body,
  });
}

function makeEnv(options = {}) {
  const messages = [];
  const emailCalls = [];
  const rateLimitKeys = [];
  return {
    messages,
    emailCalls,
    rateLimitKeys,
    env: {
      EMAIL: {
        async send(message) {
          emailCalls.push(message);
          const indexedError = options.emailErrors?.[emailCalls.length - 1];
          if (indexedError) throw indexedError;
          if (options.emailError) throw options.emailError;
          messages.push(message);
          return { messageId: "test-message-id" };
        },
      },
      RESET_RATE_LIMITER: {
        async limit({ key }) {
          if (options.rateLimitError) throw options.rateLimitError;
          rateLimitKeys.push(key);
          return { success: options.rateLimitSuccess ?? true };
        },
      },
    },
  };
}

async function responseJson(response) {
  return JSON.parse(await response.text());
}

test("sends the fixed password reset email with the pinned fragment URL", async () => {
  const { env, messages } = makeEnv();
  const response = await handleRequest(
    request(
      "/password-reset",
      JSON.stringify({ email: "coach+fall@example.com", token: RESET_TOKEN })
    ),
    env
  );

  assert.equal(response.status, 200);
  assert.deepEqual(await responseJson(response), { ok: true });
  assert.equal(response.headers.get("cache-control"), "no-store");
  assert.equal(messages.length, 1);

  const message = messages[0];
  const expectedUrl =
    "https://playbook-generator.pages.dev/reset-password" +
    `#email=coach%2Bfall%40example.com&token=${RESET_TOKEN}`;
  assert.equal(message.to, "coach+fall@example.com");
  assert.deepEqual(message.from, {
    email: "no-reply@greenwichsportssystems.com",
    name: "GSS Playbook Editor",
  });
  assert.equal(message.subject, "Reset your GSS Playbook Editor password");
  assert.match(message.text, new RegExp(expectedUrl.replace(/[.*+?^${}()|[\]\\]/g, "\\$&")));
  assert.ok(message.html.includes(expectedUrl.replace("&", "&amp;")));
  assert.ok(!message.html.includes("coach+fall@example.com"));
});

test("rejects malformed or expanded password-reset payloads without sending", async () => {
  const invalidBodies = [
    { email: "not-an-email", token: RESET_TOKEN },
    { email: "coach@example.com", token: "A".repeat(42) },
    { email: "coach@example.com", token: `${RESET_TOKEN}!` },
    { email: "coach@example.com", token: RESET_TOKEN, subject: "Injected" },
  ];

  for (const body of invalidBodies) {
    const { env, messages } = makeEnv();
    const response = await handleRequest(
      request("/password-reset", JSON.stringify(body)),
      env
    );
    assert.equal(response.status, 400);
    assert.equal(messages.length, 0);
  }
});

test("accepts every email shape accepted by the Pages account validator", async () => {
  const acceptedEmails = [
    "coach@bücher.de",
    "coach@foo_bar.com",
    '"coach"@example.com',
    `${"a".repeat(65)}@example.com`,
  ];

  for (const email of acceptedEmails) {
    const { env, messages } = makeEnv();
    const response = await handleRequest(
      request("/password-reset", JSON.stringify({ email, token: RESET_TOKEN })),
      env
    );
    assert.equal(response.status, 200);
    assert.equal(messages.length, 1);
    assert.equal(messages[0].to, email);
  }
});

test("enforces content type, valid JSON, and the request-body byte limit", async () => {
  const { env, messages } = makeEnv();

  const wrongType = await handleRequest(
    request("/password-reset", "{}", { omitContentType: true }),
    env
  );
  assert.equal(wrongType.status, 415);

  const invalidJson = await handleRequest(request("/password-reset", "{"), env);
  assert.equal(invalidJson.status, 400);

  const oversized = await handleRequest(
    request("/password-reset", JSON.stringify({ value: "x".repeat(1100) })),
    env
  );
  assert.equal(oversized.status, 413);
  assert.equal(messages.length, 0);
});

test("exposes only the two exact POST paths", async () => {
  const { env } = makeEnv();
  const missing = await handleRequest(request("/", "{}"), env);
  assert.equal(missing.status, 404);

  const trailingSlash = await handleRequest(request("/password-reset/", "{}"), env);
  assert.equal(trailingSlash.status, 404);

  const get = await handleRequest(
    request("/password-reset", undefined, { method: "GET" }),
    env
  );
  assert.equal(get.status, 405);
  assert.equal(get.headers.get("allow"), "POST");
});

test("retries transient email errors with bounded backoff and preserves the message", async () => {
  const firstError = Object.assign(new Error("temporary delivery failure"), {
    code: "E_DELIVERY_FAILED",
  });
  const secondError = Object.assign(new Error("temporary internal failure"), {
    code: "E_INTERNAL_SERVER_ERROR",
  });
  const { env, messages, emailCalls } = makeEnv({
    emailErrors: [firstError, secondError],
  });
  const delays = [];
  const response = await handleRequest(
    request(
      "/password-reset",
      JSON.stringify({ email: "coach@example.com", token: RESET_TOKEN })
    ),
    env,
    async (milliseconds) => {
      delays.push(milliseconds);
    }
  );

  assert.equal(response.status, 200);
  assert.deepEqual(await responseJson(response), { ok: true });
  assert.deepEqual(delays, [100, 300]);
  assert.equal(emailCalls.length, 3);
  assert.equal(messages.length, 1);
  assert.strictEqual(emailCalls[0], emailCalls[1]);
  assert.strictEqual(emailCalls[1], emailCalls[2]);
});

test("returns a sanitized 502 after the bounded transient retry budget", async () => {
  const error = Object.assign(new Error("contains-sensitive-provider-detail"), {
    code: "E_RATE_LIMIT_EXCEEDED",
  });
  const { env, emailCalls } = makeEnv({ emailError: error });
  const delays = [];
  const originalConsoleError = console.error;
  console.error = () => {};
  try {
    const response = await handleRequest(
      request(
        "/password-reset",
        JSON.stringify({ email: "coach@example.com", token: RESET_TOKEN })
      ),
      env,
      async (milliseconds) => {
        delays.push(milliseconds);
      }
    );
    assert.equal(response.status, 502);
    assert.deepEqual(await responseJson(response), {
      error: "Unable to send password reset email",
    });
    assert.deepEqual(delays, [100, 300]);
    assert.equal(emailCalls.length, 3);
  } finally {
    console.error = originalConsoleError;
  }
});

test("does not retry permanent Email Service errors", async () => {
  const error = Object.assign(new Error("sender is not ready"), {
    code: "E_SENDER_NOT_VERIFIED",
  });
  const { env, emailCalls } = makeEnv({ emailError: error });
  const delays = [];
  const originalConsoleError = console.error;
  console.error = () => {};
  try {
    const response = await handleRequest(
      request(
        "/password-reset",
        JSON.stringify({ email: "coach@example.com", token: RESET_TOKEN })
      ),
      env,
      async (milliseconds) => {
        delays.push(milliseconds);
      }
    );
    assert.equal(response.status, 502);
    assert.deepEqual(delays, []);
    assert.equal(emailCalls.length, 1);
  } finally {
    console.error = originalConsoleError;
  }
});

test("checks the configured rate-limit binding and returns its decision", async () => {
  for (const allowed of [true, false]) {
    const { env, rateLimitKeys } = makeEnv({ rateLimitSuccess: allowed });
    const response = await handleRequest(
      request("/password-reset/permit", JSON.stringify({ key: RATE_LIMIT_KEY })),
      env
    );
    assert.equal(response.status, 200);
    assert.deepEqual(await responseJson(response), { allowed });
    assert.deepEqual(rateLimitKeys, [RATE_LIMIT_KEY]);
  }
});

test("requires an exact lowercase SHA-256 rate-limit key payload", async () => {
  const invalidBodies = [
    { key: "A".repeat(64) },
    { key: "a".repeat(63) },
    { key: RATE_LIMIT_KEY, extra: true },
    {},
  ];

  for (const body of invalidBodies) {
    const { env, rateLimitKeys } = makeEnv();
    const response = await handleRequest(
      request("/password-reset/permit", JSON.stringify(body)),
      env
    );
    assert.equal(response.status, 400);
    assert.deepEqual(rateLimitKeys, []);
  }
});

test("fails the permit check closed when the rate limiter is unavailable", async () => {
  const { env } = makeEnv({ rateLimitError: new Error("binding unavailable") });
  const originalConsoleError = console.error;
  console.error = () => {};
  try {
    const response = await handleRequest(
      request("/password-reset/permit", JSON.stringify({ key: RATE_LIMIT_KEY })),
      env
    );
    assert.equal(response.status, 503);
    assert.deepEqual(await responseJson(response), { allowed: false });
  } finally {
    console.error = originalConsoleError;
  }
});
