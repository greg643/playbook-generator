import {
  jsonNoStore,
  normalizeEmail,
  isValidEmail,
  emailKey,
  hashPassword,
  createSessionCookie,
  createRecoveryFields,
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
    const password = body.password;
    if (!isValidEmail(email)) {
      return jsonNoStore({ error: "Invalid email address" }, { status: 400 });
    }
    if (typeof password !== "string" || password.length < 8) {
      return jsonNoStore({ error: "Password must be at least 8 characters" }, { status: 400 });
    }

    const key = await emailKey(email);
    const existing = await env.PLAYBOOK_BUCKET.get(key);
    if (existing) {
      return jsonNoStore({ error: "Account already exists" }, { status: 409 });
    }

    const saltBytes = new Uint8Array(16);
    crypto.getRandomValues(saltBytes);
    let salt = "";
    for (const b of saltBytes) salt += b.toString(16).padStart(2, "0");

    const { recoveryCode, fields: recoveryFields } = await createRecoveryFields();

    const userId = crypto.randomUUID();
    const record = {
      userId,
      email,
      salt,
      iterations: PBKDF2_ITERATIONS,
      hash: await hashPassword(password, salt, PBKDF2_ITERATIONS),
      ...recoveryFields,
      createdAt: new Date().toISOString(),
    };

    await env.PLAYBOOK_BUCKET.put(key, JSON.stringify(record), {
      httpMetadata: { contentType: "application/json" },
    });

    const cookie = await createSessionCookie(userId, email, env);
    return jsonNoStore({ email, recoveryCode }, { headers: { "Set-Cookie": cookie } });
  } catch (err) {
    console.error("Register error:", err);
    return jsonNoStore({ error: "Internal server error" }, { status: 500 });
  }
}
