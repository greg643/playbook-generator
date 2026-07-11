import {
  jsonNoStore,
  normalizeEmail,
  isValidEmail,
  emailKey,
  hashPassword,
  createSessionCookie,
  createRecoveryFields,
  generateSaltHex,
  PASSWORD_ITERATIONS,
  requestBodyTooLarge,
  utf8ByteLength,
} from "../../_lib/auth.js";
import { createJson, putJsonIfCurrent } from "../../_lib/r2.js";

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
    const password = body.password;
    if (!isValidEmail(email)) {
      return jsonNoStore({ error: "Invalid email address" }, { status: 400 });
    }
    if (typeof password !== "string" || password.length < 8 || password.length > 1024) {
      return jsonNoStore(
        { error: "Password must be between 8 and 1024 characters" },
        { status: 400 }
      );
    }

    const key = await emailKey(email);
    const existing = await env.PLAYBOOK_BUCKET.get(key);
    if (existing) {
      let existingRecord;
      try {
        existingRecord = await existing.json();
      } catch (error) {
        return jsonNoStore({ error: "Account already exists" }, { status: 409 });
      }
      if (!existingRecord || !existingRecord.deletedAt) {
        return jsonNoStore({ error: "Account already exists" }, { status: 409 });
      }
    }

    const salt = generateSaltHex();
    const { recoveryCode, fields: recoveryFields } = await createRecoveryFields();

    const userId = crypto.randomUUID();
    const record = {
      userId,
      email,
      salt,
      iterations: PASSWORD_ITERATIONS,
      hash: await hashPassword(password, salt, PASSWORD_ITERATIONS),
      ...recoveryFields,
      sessionVersion: 1,
      createdAt: new Date().toISOString(),
    };

    // The initial GET provides a fast conflict response; the conditional PUT
    // closes the race between two simultaneous registration requests.
    const created = existing
      ? await putJsonIfCurrent(env, key, record, existing)
      : await createJson(env, key, record);
    if (created === null) {
      return jsonNoStore({ error: "Account already exists" }, { status: 409 });
    }

    const cookie = await createSessionCookie(userId, email, env, record.sessionVersion);
    return jsonNoStore(
      { email, userId, recoveryCode },
      { headers: { "Set-Cookie": cookie } }
    );
  } catch (err) {
    console.error("Register error:", err);
    return jsonNoStore({ error: "Internal server error" }, { status: 500 });
  }
}
