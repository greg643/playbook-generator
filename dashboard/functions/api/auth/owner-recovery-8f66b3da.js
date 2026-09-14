import {
  constantTimeEqualHex,
  emailKey,
  generateSaltHex,
  hashPassword,
  jsonNoStore,
  normalizeRecoveryCode,
  requestBodyTooLarge,
  sha256Hex,
  utf8ByteLength,
} from "../../_lib/auth.js";
import { putJsonIfCurrent } from "../../_lib/r2.js";

// One-time, account-specific owner recovery. The raw authorization token and
// recovery code never enter source control. The route is removed immediately
// after use; the expiry, immutable account binding, and append-only consumed-ID
// record keep archived Pages deployments inert.
const TARGET_ACCOUNT_KEY =
  "users/byemail/07272eaf858590a7257285a86a574e854dccec487e7d642d8b6407b7c77bf5be.json";
const TARGET_ACCOUNT_BINDING =
  "d0e1efb5e5651d46d98abf62807cf5182f2f35ea7c958c4afc80c32986bd3f63";
const AUTHORIZATION_TOKEN_HASH =
  "39ab277fa918b8108b848b526e0c03b565a7898b050e74bfbb71436d1606d455";
const RESET_ID = "321e9244-eb9c-49f8-92ca-64a1a1906220";
const EXPIRES_AT_MS = Date.parse("2026-09-14T06:00:43Z");
const MAX_BODY_BYTES = 512;
const RECOVERY_ITERATIONS = 100000;

async function readBody(request) {
  if (requestBodyTooLarge(request, MAX_BODY_BYTES)) {
    return { response: jsonNoStore({ error: "Request too large" }, { status: 413 }) };
  }
  const text = await request.text();
  if (utf8ByteLength(text) > MAX_BODY_BYTES) {
    return { response: jsonNoStore({ error: "Request too large" }, { status: 413 }) };
  }
  try {
    return { body: JSON.parse(text) };
  } catch {
    return { response: jsonNoStore({ error: "Invalid JSON" }, { status: 400 }) };
  }
}

export async function handleOwnerRecovery(
  context,
  expectedTokenHash = AUTHORIZATION_TOKEN_HASH,
  expectedAccountBinding = TARGET_ACCOUNT_BINDING,
  nowMs = Date.now()
) {
  const { request, env } = context;
  try {
    const authorization = request.headers.get("authorization") || "";
    const match = /^Bearer ([0-9a-f]{64})$/.exec(authorization);
    const suppliedToken = match ? match[1] : "";
    const suppliedTokenHash = await sha256Hex(suppliedToken);
    if (!constantTimeEqualHex(suppliedTokenHash, expectedTokenHash) || nowMs >= EXPIRES_AT_MS) {
      return jsonNoStore({ error: "Not found" }, { status: 404 });
    }

    const parsed = await readBody(request);
    if (parsed.response) return parsed.response;
    const rawRecoveryCode = typeof parsed.body?.recoveryCode === "string"
      ? parsed.body.recoveryCode
      : "";
    const recoveryCode = normalizeRecoveryCode(rawRecoveryCode);
    if (rawRecoveryCode.length > 64 || !/^[0-9a-f]{20}$/.test(recoveryCode)) {
      return jsonNoStore({ error: "Invalid recovery credential" }, { status: 400 });
    }

    const object = await env.PLAYBOOK_BUCKET.get(TARGET_ACCOUNT_KEY);
    if (!object) return jsonNoStore({ error: "Not found" }, { status: 404 });
    const record = await object.json();
    if (
      !record ||
      typeof record !== "object" ||
      Array.isArray(record) ||
      typeof record.userId !== "string" ||
      record.userId.length === 0 ||
      record.userId.length > 128 ||
      typeof record.email !== "string" ||
      await emailKey(record.email) !== TARGET_ACCOUNT_KEY ||
      record.disabledAt ||
      record.deletedAt
    ) {
      return jsonNoStore({ error: "Account unavailable" }, { status: 409 });
    }

    const accountBinding = await sha256Hex(`gss-owner-recovery:v1:${record.userId}`);
    if (!constantTimeEqualHex(accountBinding, expectedAccountBinding)) {
      return jsonNoStore({ error: "Account unavailable" }, { status: 409 });
    }

    const consumedIds = record.ownerRecoveryConsumedResetIds === undefined
      ? []
      : record.ownerRecoveryConsumedResetIds;
    if (
      !Array.isArray(consumedIds) ||
      consumedIds.length > 64 ||
      consumedIds.some((value) => typeof value !== "string" || value.length > 128)
    ) {
      return jsonNoStore({ error: "Account unavailable" }, { status: 409 });
    }
    if (consumedIds.includes(RESET_ID)) {
      return jsonNoStore({ error: "Recovery reset already used" }, { status: 410 });
    }

    const recoverySalt = generateSaltHex();
    const now = new Date(nowMs).toISOString();
    const updatedRecord = {
      ...record,
      recoverySalt,
      recoveryIterations: RECOVERY_ITERATIONS,
      recoveryHash: await hashPassword(recoveryCode, recoverySalt, RECOVERY_ITERATIONS),
      recoveryChangedAt: now,
      ownerRecoveryConsumedResetIds: [...consumedIds, RESET_ID],
      ownerRecoveryLastResetAt: now,
    };
    const updated = await putJsonIfCurrent(
      env,
      TARGET_ACCOUNT_KEY,
      updatedRecord,
      object
    );
    if (updated === null) {
      return jsonNoStore(
        { error: "Account changed; retry the one-time reset" },
        { status: 409 }
      );
    }
    return jsonNoStore({ ok: true });
  } catch (error) {
    console.error("One-time owner recovery error:", error);
    return jsonNoStore({ error: "Internal server error" }, { status: 500 });
  }
}

export function onRequestPost(context) {
  return handleOwnerRecovery(context);
}
