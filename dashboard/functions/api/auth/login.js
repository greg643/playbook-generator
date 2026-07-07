import {
  jsonNoStore,
  normalizeEmail,
  isValidEmail,
  emailKey,
  hashPassword,
  createSessionCookie,
} from "../../_lib/auth.js";

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
    if (!isValidEmail(email) || typeof password !== "string") {
      return jsonNoStore({ error: "Invalid email or password" }, { status: 401 });
    }

    const key = await emailKey(email);
    const obj = await env.PLAYBOOK_BUCKET.get(key);
    if (!obj) {
      return jsonNoStore({ error: "Invalid email or password" }, { status: 401 });
    }

    const record = await obj.json();
    const hash = await hashPassword(password, record.salt, record.iterations);

    // Constant-time compare of hex digests
    let diff = hash.length === record.hash.length ? 0 : 1;
    for (let i = 0; i < hash.length; i++) {
      diff |= hash.charCodeAt(i) ^ record.hash.charCodeAt(i % record.hash.length);
    }
    if (diff !== 0) {
      return jsonNoStore({ error: "Invalid email or password" }, { status: 401 });
    }

    const cookie = await createSessionCookie(record.userId, record.email, env);
    return jsonNoStore({ email: record.email }, { headers: { "Set-Cookie": cookie } });
  } catch (err) {
    console.error("Login error:", err);
    return jsonNoStore({ error: "Internal server error" }, { status: 500 });
  }
}
