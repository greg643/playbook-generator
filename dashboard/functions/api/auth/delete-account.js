import {
  jsonNoStore,
  requireUser,
  emailKey,
  hashPassword,
  constantTimeEqualHex,
  clearSessionCookie,
} from "../../_lib/auth.js";

export async function onRequestPost(context) {
  const { request, env } = context;

  try {
    const { user, response } = await requireUser(context);
    if (!user) return response;

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

    const password = body.password;
    if (typeof password !== "string") {
      return jsonNoStore({ error: "Invalid password" }, { status: 401 });
    }

    const key = await emailKey(user.email);
    const obj = await env.PLAYBOOK_BUCKET.get(key);
    if (!obj) {
      return jsonNoStore({ error: "Invalid password" }, { status: 401 });
    }

    const record = await obj.json();
    const hash = await hashPassword(password, record.salt, record.iterations);
    if (!constantTimeEqualHex(hash, record.hash)) {
      return jsonNoStore({ error: "Invalid password" }, { status: 401 });
    }

    await env.PLAYBOOK_BUCKET.delete(key);
    await env.PLAYBOOK_BUCKET.delete(`accounts/${record.userId}/playbook.json`);

    return jsonNoStore({ ok: true }, { headers: { "Set-Cookie": clearSessionCookie() } });
  } catch (err) {
    console.error("Delete-account error:", err);
    return jsonNoStore({ error: "Internal server error" }, { status: 500 });
  }
}
