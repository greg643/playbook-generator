import { jsonNoStore, requireUser } from "../_lib/auth.js";

const MAX_DOC_BYTES = 1000000;
const MAX_OFFENSE = 16;
const MAX_DEFENSE = 6;
const MAX_NAME_LENGTH = 60;

function playbookKey(userId) {
  return `accounts/${userId}/playbook.json`;
}

export async function onRequestGet(context) {
  const { env } = context;

  try {
    const { user, response } = await requireUser(context);
    if (!user) return response;

    const obj = await env.PLAYBOOK_BUCKET.get(playbookKey(user.userId));
    if (!obj) {
      return jsonNoStore({ schema: 1, offense: [], defense: [] });
    }
    const text = await obj.text();
    return new Response(text, {
      headers: { "Content-Type": "application/json", "Cache-Control": "no-store" },
    });
  } catch (err) {
    console.error("Plays GET error:", err);
    return jsonNoStore({ error: "Internal server error" }, { status: 500 });
  }
}

function validatePlays(list, max) {
  if (!Array.isArray(list) || list.length > max) return false;
  for (const play of list) {
    if (!play || typeof play !== "object" || Array.isArray(play)) return false;
    if (typeof play.name !== "string" || play.name.length > MAX_NAME_LENGTH) return false;
    if (!play.chips || typeof play.chips !== "object" || Array.isArray(play.chips)) return false;
  }
  return true;
}

export async function onRequestPut(context) {
  const { request, env } = context;

  try {
    const { user, response } = await requireUser(context);
    if (!user) return response;

    const bodyText = await request.text();
    if (new TextEncoder().encode(bodyText).length > MAX_DOC_BYTES) {
      return jsonNoStore({ error: "Playbook too large (max 1 MB)" }, { status: 400 });
    }

    let doc;
    try {
      doc = JSON.parse(bodyText);
    } catch (err) {
      return jsonNoStore({ error: "Invalid JSON" }, { status: 400 });
    }

    if (
      !doc ||
      typeof doc !== "object" ||
      Array.isArray(doc) ||
      !validatePlays(doc.offense, MAX_OFFENSE) ||
      !validatePlays(doc.defense, MAX_DEFENSE)
    ) {
      return jsonNoStore({ error: "Invalid playbook document" }, { status: 400 });
    }

    doc.updatedAt = new Date().toISOString();
    await env.PLAYBOOK_BUCKET.put(playbookKey(user.userId), JSON.stringify(doc), {
      httpMetadata: { contentType: "application/json" },
    });

    return jsonNoStore({ ok: true });
  } catch (err) {
    console.error("Plays PUT error:", err);
    return jsonNoStore({ error: "Internal server error" }, { status: 500 });
  }
}
