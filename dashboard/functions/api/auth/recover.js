import {
  jsonNoStore,
  normalizeEmail,
  isValidEmail,
  emailKey,
  hashPassword,
  createSessionCookie,
  createRecoveryFields,
  normalizeRecoveryCode,
  constantTimeEqualHex,
  generateSaltHex,
  PASSWORD_ITERATIONS,
  requestBodyTooLarge,
  utf8ByteLength,
} from "../../_lib/auth.js";
import { putJsonIfCurrent } from "../../_lib/r2.js";

const MAX_BODY_BYTES = 10000;

export async function onRequestPost(context) {
  const { request, env } = context;

  try {
    if (requestBodyTooLarge(request, MAX_BODY_BYTES)) {
      return jsonNoStore({ error: "Request too large" }, { status: 413 });
    }
    const bodyText = await request.text();
    if (utf8ByteLength(bodyText) > MAX_BODY_BYTES) {
      return jsonNoStore({ error: "Request too large" }, { status: 413 });
    }

    let body;
    try {
      body = JSON.parse(bodyText);
    } catch (err) {
      return jsonNoStore({ error: "Invalid JSON" }, { status: 400 });
    }

    const email = normalizeEmail(body.email);
    const newPassword = body.newPassword;
    if (!isValidEmail(email)) {
      return jsonNoStore({ error: "Invalid email address" }, { status: 400 });
    }
    if (typeof newPassword !== "string" || newPassword.length < 8 || newPassword.length > 1024) {
      return jsonNoStore(
        { error: "Password must be between 8 and 1024 characters" },
        { status: 400 }
      );
    }

    const key = await emailKey(email);
    const obj = await env.PLAYBOOK_BUCKET.get(key);
    if (!obj) {
      return jsonNoStore({ error: "Invalid email or recovery code" }, { status: 401 });
    }

    const record = await obj.json();
    const code = normalizeRecoveryCode(body.recoveryCode);
    if (
      code.length !== 20 ||
      !record ||
      typeof record.userId !== "string" ||
      normalizeEmail(record.email) !== email ||
      record.disabledAt ||
      typeof record.recoveryHash !== "string" ||
      typeof record.recoverySalt !== "string" ||
      typeof record.recoveryIterations !== "number"
    ) {
      return jsonNoStore({ error: "Invalid email or recovery code" }, { status: 401 });
    }

    const codeHash = await hashPassword(code, record.recoverySalt, record.recoveryIterations);
    if (!constantTimeEqualHex(codeHash, record.recoveryHash)) {
      return jsonNoStore({ error: "Invalid email or recovery code" }, { status: 401 });
    }

    // Set the new password and rotate the recovery code.
    record.salt = generateSaltHex();
    record.iterations = PASSWORD_ITERATIONS;
    record.hash = await hashPassword(newPassword, record.salt, record.iterations);
    const { recoveryCode, fields } = await createRecoveryFields();
    Object.assign(record, fields);
    record.sessionVersion =
      (Number.isSafeInteger(record.sessionVersion) && record.sessionVersion >= 1
        ? record.sessionVersion
        : 1) + 1;
    record.passwordChangedAt = new Date().toISOString();

    // Consume the recovery code exactly once. A simultaneous recovery request
    // using the same code loses the ETag precondition and cannot overwrite the
    // first password change.
    const updated = await putJsonIfCurrent(env, key, record, obj);
    if (updated === null) {
      return jsonNoStore(
        { error: "Recovery code was already used; try the newly issued code" },
        { status: 409 }
      );
    }

    const cookie = await createSessionCookie(
      record.userId,
      record.email,
      env,
      record.sessionVersion
    );
    return jsonNoStore(
      { email: record.email, userId: record.userId, recoveryCode },
      { headers: { "Set-Cookie": cookie } }
    );
  } catch (err) {
    console.error("Recover error:", err);
    return jsonNoStore({ error: "Internal server error" }, { status: 500 });
  }
}
