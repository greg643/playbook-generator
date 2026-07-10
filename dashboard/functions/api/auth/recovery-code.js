import { jsonNoStore, requireUser, emailKey, createRecoveryFields } from "../../_lib/auth.js";

export async function onRequestPost(context) {
  const { env } = context;

  try {
    const { user, response } = await requireUser(context);
    if (!user) return response;

    const key = await emailKey(user.email);
    const obj = await env.PLAYBOOK_BUCKET.get(key);
    if (!obj) {
      // Valid session but the account record is gone (e.g. deleted).
      return jsonNoStore({ error: "Not signed in" }, { status: 401 });
    }

    const record = await obj.json();
    const { recoveryCode, fields } = await createRecoveryFields();
    Object.assign(record, fields);

    await env.PLAYBOOK_BUCKET.put(key, JSON.stringify(record), {
      httpMetadata: { contentType: "application/json" },
    });

    return jsonNoStore({ recoveryCode });
  } catch (err) {
    console.error("Recovery-code error:", err);
    return jsonNoStore({ error: "Internal server error" }, { status: 500 });
  }
}
