import {
  jsonNoStore,
  isAccountActive,
  requestBodyTooLarge,
  requireUser,
  utf8ByteLength,
} from "../_lib/auth.js";
import { putJsonIfCurrent } from "../_lib/r2.js";

const MAX_DOC_BYTES = 1000000;
// Storage caps (an archive/"bench" beyond what one PDF holds); the generate
// flow separately enforces at most 16 offense / 6 defense INCLUDED plays.
const MAX_OFFENSE = 64;
const MAX_DEFENSE = 24;
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

    if (requestBodyTooLarge(request, MAX_DOC_BYTES)) {
      return jsonNoStore({ error: "Playbook too large (max 1 MB)" }, { status: 413 });
    }
    const bodyText = await request.text();
    if (utf8ByteLength(bodyText) > MAX_DOC_BYTES) {
      return jsonNoStore({ error: "Playbook too large (max 1 MB)" }, { status: 413 });
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

    // Bind the browser's in-memory document to the authenticated account. A
    // session can change while an old autosave is queued (for example after
    // expiry followed by sign-in to a different account); never let that old
    // body land in the new account.
    if (doc.ownerId !== user.userId) {
      return jsonNoStore({ error: "Playbook account changed" }, { status: 403 });
    }
    delete doc.ownerId;

    // Optional optimistic-concurrency check: the editor sends the updatedAt it
    // loaded as baseUpdatedAt; if the stored doc has moved on, reject with 409.
    // Absent baseUpdatedAt means save unconditionally (back-compat/overwrite).
    const baseUpdatedAt = doc.baseUpdatedAt;
    delete doc.baseUpdatedAt;
    const existing = await env.PLAYBOOK_BUCKET.get(playbookKey(user.userId));
    if (baseUpdatedAt !== undefined) {
      if (existing) {
        const stored = await existing.json();
        if (stored && stored.updatedAt && stored.updatedAt !== baseUpdatedAt) {
          return jsonNoStore(
            { error: "conflict", serverUpdatedAt: stored.updatedAt },
            { status: 409 }
          );
        }
      } else if (baseUpdatedAt !== null && baseUpdatedAt !== "") {
        return jsonNoStore(
          { error: "conflict", serverUpdatedAt: null },
          { status: 409 }
        );
      }
    }

    doc.updatedAt = new Date().toISOString();
    const saved = await putJsonIfCurrent(
      env,
      playbookKey(user.userId),
      doc,
      existing
    );
    if (saved === null) {
      const latest = await env.PLAYBOOK_BUCKET.get(playbookKey(user.userId));
      let serverUpdatedAt = null;
      if (latest) {
        try {
          serverUpdatedAt = (await latest.json()).updatedAt || null;
        } catch (err) {
          console.error("Could not read conflicting playbook:", err);
        }
      }
      return jsonNoStore({ error: "conflict", serverUpdatedAt }, { status: 409 });
    }

    // Close the race with account deletion: either this check happens before
    // the deletion sweep (which then removes the playbook), or it observes the
    // tombstone and removes the just-finished save itself.
    if (!(await isAccountActive(env, user))) {
      await env.PLAYBOOK_BUCKET.delete(playbookKey(user.userId));
      return jsonNoStore({ error: "Account is being deleted" }, { status: 409 });
    }

    return jsonNoStore({ ok: true, updatedAt: doc.updatedAt });
  } catch (err) {
    console.error("Plays PUT error:", err);
    return jsonNoStore({ error: "Internal server error" }, { status: 500 });
  }
}
