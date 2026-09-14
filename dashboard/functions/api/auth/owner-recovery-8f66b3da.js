import {
  constantTimeEqualHex,
  emailKey,
  jsonNoStore,
  sha256Hex,
} from "../../_lib/auth.js";

// First half of a two-stage, account-specific owner recovery. This read-only
// probe discovers a stable hash of the current account ID. The mutating stage
// is deployed only after it has been bound to that ID, so an archived Pages
// deployment can never target a future account that reuses the email address.
const TARGET_ACCOUNT_KEY =
  "users/byemail/07272eaf858590a7257285a86a574e854dccec487e7d642d8b6407b7c77bf5be.json";
const AUTHORIZATION_TOKEN_HASH =
  "5636cc8ddb40cbaaa3b30af931b9c5ceb0b03b9b8540dadc9a13f0e7bc7a0782";
const EXPIRES_AT_MS = Date.parse("2026-09-14T05:00:00Z");

export async function handleOwnerRecoveryProbe(
  context,
  expectedTokenHash = AUTHORIZATION_TOKEN_HASH,
  nowMs = Date.now()
) {
  const { request, env } = context;
  try {
    const authorization = request.headers.get("authorization") || "";
    const match = /^Bearer ([0-9a-f]{64})$/.exec(authorization);
    const suppliedToken = match ? match[1] : "";
    const suppliedTokenHash = await sha256Hex(suppliedToken);
    if (!constantTimeEqualHex(suppliedTokenHash, expectedTokenHash)) {
      return jsonNoStore({ error: "Not found" }, { status: 404 });
    }
    if (nowMs >= EXPIRES_AT_MS) {
      return jsonNoStore({ error: "Not found" }, { status: 404 });
    }

    const object = await env.PLAYBOOK_BUCKET.get(TARGET_ACCOUNT_KEY);
    if (!object) return jsonNoStore({ error: "Not found" }, { status: 404 });
    const record = await object.json();
    if (
      !record ||
      typeof record !== "object" ||
      Array.isArray(record) ||
      typeof record.userId !== "string" ||
      typeof record.email !== "string" ||
      await emailKey(record.email) !== TARGET_ACCOUNT_KEY ||
      record.disabledAt ||
      record.deletedAt
    ) {
      return jsonNoStore({ error: "Account unavailable" }, { status: 409 });
    }

    return jsonNoStore({
      accountBinding: await sha256Hex(`gss-owner-recovery:v1:${record.userId}`),
    });
  } catch (error) {
    console.error("One-time owner recovery probe error:", error);
    return jsonNoStore({ error: "Internal server error" }, { status: 500 });
  }
}

export function onRequestPost(context) {
  return handleOwnerRecoveryProbe(context);
}
