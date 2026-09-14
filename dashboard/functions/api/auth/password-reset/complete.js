import {
  clearSessionCookie,
  constantTimeEqualHex,
  emailKey,
  hashPasswordResetToken,
  isValidEmail,
  isValidPasswordResetToken,
  jsonNoStore,
  normalizeEmail,
  readBoundedUtf8Text,
  replacePasswordAndRecovery,
} from "../../../_lib/auth.js";
import { isUsablePasswordResetChallenge } from "../../../_lib/password-reset.js";
import { putJsonIfCurrent } from "../../../_lib/r2.js";

const MAX_BODY_BYTES = 10000;
const INVALID_LINK_MESSAGE = "This password-reset link is invalid or has expired.";

function invalidLinkResponse(status = 401) {
  return jsonNoStore({ error: INVALID_LINK_MESSAGE }, { status });
}

function recordAcceptsToken(record, email, tokenHash, nowMs) {
  return Boolean(
    record &&
      normalizeEmail(record.email) === email &&
      !record.deletedAt &&
      !record.disabledAt &&
      isUsablePasswordResetChallenge(record, nowMs) &&
      constantTimeEqualHex(tokenHash, record.passwordReset.tokenHash)
  );
}

async function committedAfterAmbiguousWrite(env, key, expectedRecord, requestId) {
  const object = await env.PLAYBOOK_BUCKET.get(key);
  if (!object) return false;
  try {
    const record = await object.json();
    return Boolean(
      record &&
        record.userId === expectedRecord.userId &&
        record.lastPasswordResetId === requestId &&
        record.sessionVersion === expectedRecord.sessionVersion &&
        record.hash === expectedRecord.hash &&
        record.recoveryHash === expectedRecord.recoveryHash &&
        !record.passwordReset
    );
  } catch (error) {
    return false;
  }
}

export async function onRequestPost(context) {
  const { request, env } = context;

  try {
    const bodyText = await readBoundedUtf8Text(request, MAX_BODY_BYTES);
    if (bodyText === null) {
      return jsonNoStore({ error: "Request too large" }, { status: 413 });
    }

    let body;
    try {
      body = JSON.parse(bodyText);
    } catch (error) {
      return jsonNoStore({ error: "Invalid JSON" }, { status: 400 });
    }

    const email = normalizeEmail(body && body.email);
    const token = body && body.token;
    const newPassword = body && body.newPassword;
    if (!isValidEmail(email) || !isValidPasswordResetToken(token)) {
      return invalidLinkResponse();
    }
    if (typeof newPassword !== "string" || newPassword.length < 8 || newPassword.length > 1024) {
      return jsonNoStore(
        { error: "Password must be between 8 and 1024 characters" },
        { status: 400 }
      );
    }

    const key = await emailKey(email);
    let object = await env.PLAYBOOK_BUCKET.get(key);
    if (!object) return invalidLinkResponse();
    let record;
    try {
      record = await object.json();
    } catch (error) {
      return invalidLinkResponse();
    }

    const nowMs = Date.now();
    const tokenHash = await hashPasswordResetToken(token);
    if (!recordAcceptsToken(record, email, tokenHash, nowMs)) {
      return invalidLinkResponse();
    }
    const requestId = record.passwordReset.requestId;

    // The KDFs are intentionally expensive. Re-read the credential object
    // afterward so a password change, deletion, or newer reset request that
    // happened during that work wins before we attempt the final CAS.
    const replacement = await replacePasswordAndRecovery(record, newPassword);
    const latestObject = await env.PLAYBOOK_BUCKET.get(key);
    if (!latestObject) return invalidLinkResponse();
    let latestRecord;
    try {
      latestRecord = await latestObject.json();
    } catch (error) {
      return invalidLinkResponse();
    }
    if (!recordAcceptsToken(latestRecord, email, tokenHash, Date.now())) {
      return invalidLinkResponse(409);
    }

    const updatedRecord = {
      ...latestRecord,
      salt: replacement.record.salt,
      iterations: replacement.record.iterations,
      hash: replacement.record.hash,
      recoverySalt: replacement.record.recoverySalt,
      recoveryIterations: replacement.record.recoveryIterations,
      recoveryHash: replacement.record.recoveryHash,
      sessionVersion: replacement.record.sessionVersion,
      passwordChangedAt: replacement.record.passwordChangedAt,
      lastPasswordResetId: requestId,
    };
    delete updatedRecord.passwordReset;

    let updated;
    try {
      updated = await putJsonIfCurrent(env, key, updatedRecord, latestObject);
    } catch (error) {
      if (await committedAfterAmbiguousWrite(env, key, updatedRecord, requestId)) {
        updated = true;
      } else {
        throw error;
      }
    }
    if (updated === null) return invalidLinkResponse(409);

    return jsonNoStore(
      {
        message: "Password reset. Sign in with your new password.",
        recoveryCode: replacement.recoveryCode,
      },
      { headers: { "Set-Cookie": clearSessionCookie() } }
    );
  } catch (error) {
    console.error(JSON.stringify({ event: "password_reset_completion_failed" }));
    return jsonNoStore({ error: "Internal server error" }, { status: 500 });
  }
}
