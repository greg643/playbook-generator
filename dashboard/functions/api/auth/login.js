import {
  jsonNoStore,
  normalizeEmail,
  isValidEmail,
  emailKey,
  hashPassword,
  createSessionCookie,
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
    const password = body.password;
    if (!isValidEmail(email) || typeof password !== "string" || password.length > 1024) {
      return jsonNoStore({ error: "Invalid email or password" }, { status: 401 });
    }

    const key = await emailKey(email);
    const obj = await env.PLAYBOOK_BUCKET.get(key);
    if (!obj) {
      return jsonNoStore({ error: "Invalid email or password" }, { status: 401 });
    }

    const record = await obj.json();
    if (
      !record ||
      typeof record.userId !== "string" ||
      normalizeEmail(record.email) !== email ||
      typeof record.salt !== "string" ||
      !Number.isSafeInteger(record.iterations) ||
      typeof record.hash !== "string"
    ) {
      return jsonNoStore({ error: "Invalid email or password" }, { status: 401 });
    }
    const hash = await hashPassword(password, record.salt, record.iterations);
    if (!constantTimeEqualHex(hash, record.hash)) {
      return jsonNoStore({ error: "Invalid email or password" }, { status: 401 });
    }

    if (record.disabledAt) {
      return jsonNoStore(
        {
          error: "Account deletion is pending",
          deletionPending: true,
          email: record.email,
          userId: record.userId,
        },
        { status: 423 }
      );
    }

    // Upgrade legacy password work factors after a successful login without
    // risking an overwrite of a concurrent password recovery.
    let currentRecord = record;
    if (record.iterations < PASSWORD_ITERATIONS) {
      const upgraded = {
        ...record,
        salt: generateSaltHex(),
        iterations: PASSWORD_ITERATIONS,
        sessionVersion:
          Number.isSafeInteger(record.sessionVersion) && record.sessionVersion >= 1
            ? record.sessionVersion
            : 1,
      };
      upgraded.hash = await hashPassword(password, upgraded.salt, upgraded.iterations);
      const result = await putJsonIfCurrent(env, key, upgraded, obj);
      if (result !== null) {
        currentRecord = upgraded;
      } else {
        const latestObject = await env.PLAYBOOK_BUCKET.get(key);
        if (!latestObject) {
          return jsonNoStore({ error: "Invalid email or password" }, { status: 401 });
        }
        const latest = await latestObject.json();
        const latestHash = await hashPassword(password, latest.salt, latest.iterations);
        if (
          latest.userId !== record.userId ||
          latest.disabledAt ||
          !constantTimeEqualHex(latestHash, latest.hash)
        ) {
          return jsonNoStore({ error: "Invalid email or password" }, { status: 401 });
        }
        currentRecord = latest;
      }
    }

    const sessionVersion =
      Number.isSafeInteger(currentRecord.sessionVersion) && currentRecord.sessionVersion >= 1
        ? currentRecord.sessionVersion
        : 1;
    const cookie = await createSessionCookie(
      currentRecord.userId,
      currentRecord.email,
      env,
      sessionVersion
    );
    return jsonNoStore(
      { email: currentRecord.email, userId: currentRecord.userId },
      { headers: { "Set-Cookie": cookie } }
    );
  } catch (err) {
    console.error("Login error:", err);
    return jsonNoStore({ error: "Internal server error" }, { status: 500 });
  }
}
