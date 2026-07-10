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
} from "../../_lib/auth.js";

const PBKDF2_ITERATIONS = 100000;

export async function onRequestPost(context) {
  const { request, env } = context;

  try {
    const bodyText = await request.text();
    if (bodyText.length > 10000) {
      return jsonNoStore({ error: "Request too large" }, { status: 400 });
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
    if (typeof newPassword !== "string" || newPassword.length < 8) {
      return jsonNoStore({ error: "Password must be at least 8 characters" }, { status: 400 });
    }

    const key = await emailKey(email);
    const obj = await env.PLAYBOOK_BUCKET.get(key);
    if (!obj) {
      return jsonNoStore({ error: "Invalid email or recovery code" }, { status: 401 });
    }

    const record = await obj.json();
    const code = normalizeRecoveryCode(body.recoveryCode);
    if (
      !code ||
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
    record.iterations = PBKDF2_ITERATIONS;
    record.hash = await hashPassword(newPassword, record.salt, record.iterations);
    const { recoveryCode, fields } = await createRecoveryFields();
    Object.assign(record, fields);

    await env.PLAYBOOK_BUCKET.put(key, JSON.stringify(record), {
      httpMetadata: { contentType: "application/json" },
    });

    const cookie = await createSessionCookie(record.userId, record.email, env);
    return jsonNoStore({ email: record.email, recoveryCode }, { headers: { "Set-Cookie": cookie } });
  } catch (err) {
    console.error("Recover error:", err);
    return jsonNoStore({ error: "Internal server error" }, { status: 500 });
  }
}
