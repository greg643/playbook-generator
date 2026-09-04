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
const CURRENT_SCHEMA = 2;
const SUPPORTED_PLAYER_COUNTS = new Set([5, 6]);

const CHIP_KEYS = {
  offense: {
    5: new Set(["1", "2", "3", "C", "QB"]),
    6: new Set(["1", "2", "3", "4", "5", "QB"]),
  },
  defense: {
    5: new Set(["1", "2", "3", "4", "5"]),
    6: new Set(["1", "2", "3", "4", "5", "N"]),
  },
};
const LEGACY_DEFENSE_5_CHIP_KEYS = new Set(["1", "2", "3", "4", "N"]);

function playbookKey(userId) {
  return `accounts/${userId}/playbook.json`;
}

function migrateLegacyFivePlayerDefense(doc) {
  if (
    !doc ||
    typeof doc !== "object" ||
    Array.isArray(doc) ||
    doc.schema !== CURRENT_SCHEMA ||
    !Array.isArray(doc.defense)
  ) {
    return doc;
  }

  for (const play of doc.defense) {
    if (
      !play ||
      typeof play !== "object" ||
      Array.isArray(play) ||
      play.playersPerSide !== 5 ||
      !play.chips ||
      typeof play.chips !== "object" ||
      Array.isArray(play.chips)
    ) {
      continue;
    }
    const chipKeys = Object.keys(play.chips);
    const isLegacyLineup =
      chipKeys.length === LEGACY_DEFENSE_5_CHIP_KEYS.size &&
      chipKeys.every((key) => LEGACY_DEFENSE_5_CHIP_KEYS.has(key));
    if (!isLegacyLineup) continue;

    // Preserve the stored chip object (and therefore its exact position) while
    // adopting the canonical numeric key. Routes follow the same one-way map.
    play.chips["5"] = play.chips.N;
    delete play.chips.N;
    if (Array.isArray(play.routes)) {
      for (const route of play.routes) {
        if (route && typeof route === "object" && route.chip === "N") {
          route.chip = "5";
        }
      }
    }
  }
  return doc;
}

export async function onRequestGet(context) {
  const { env } = context;

  try {
    const { user, response } = await requireUser(context);
    if (!user) return response;

    const obj = await env.PLAYBOOK_BUCKET.get(playbookKey(user.userId));
    if (!obj) {
      // null distinguishes a genuinely new account from a legacy empty
      // schema-1 playbook. The editor asks once, then persists 5 or 6.
      return jsonNoStore({
        schema: CURRENT_SCHEMA,
        defaultPlayersPerSide: null,
        offense: [],
        defense: [],
      });
    }
    const doc = migrateLegacyFivePlayerDefense(await obj.json());
    return jsonNoStore(doc);
  } catch (err) {
    console.error("Plays GET error:", err);
    return jsonNoStore({ error: "Internal server error" }, { status: 500 });
  }
}

function hasExpectedChips(chips, section, playersPerSide) {
  const expected = CHIP_KEYS[section][playersPerSide];
  const keys = Object.keys(chips);
  if (keys.length !== expected.size || keys.some((key) => !expected.has(key))) return false;
  return keys.every((key) => {
    const chip = chips[key];
    return (
      chip &&
      typeof chip === "object" &&
      !Array.isArray(chip) &&
      Number.isFinite(chip.x) &&
      Number.isFinite(chip.y) &&
      chip.x >= 0 &&
      chip.x <= 1 &&
      chip.y >= 0 &&
      chip.y <= 1
    );
  });
}

function validatePlays(list, max, section, schema) {
  if (!Array.isArray(list) || list.length > max) return false;
  for (const play of list) {
    if (!play || typeof play !== "object" || Array.isArray(play)) return false;
    if (typeof play.name !== "string" || play.name.length > MAX_NAME_LENGTH) return false;
    if (!play.chips || typeof play.chips !== "object" || Array.isArray(play.chips)) return false;
    if (schema === CURRENT_SCHEMA) {
      if (!SUPPORTED_PLAYER_COUNTS.has(play.playersPerSide)) return false;
      if (!hasExpectedChips(play.chips, section, play.playersPerSide)) return false;
      if (
        Array.isArray(play.routes) &&
        play.routes.some(
          (route) =>
            !route ||
            typeof route !== "object" ||
            typeof route.chip !== "string" ||
            !Object.hasOwn(play.chips, route.chip)
        )
      ) {
        return false;
      }
    }
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
    migrateLegacyFivePlayerDefense(doc);

    const schema = doc && doc.schema;
    if (
      !doc ||
      typeof doc !== "object" ||
      Array.isArray(doc) ||
      ![1, CURRENT_SCHEMA].includes(schema) ||
      (schema === CURRENT_SCHEMA && !SUPPORTED_PLAYER_COUNTS.has(doc.defaultPlayersPerSide)) ||
      !validatePlays(doc.offense, MAX_OFFENSE, "offense", schema) ||
      !validatePlays(doc.defense, MAX_DEFENSE, "defense", schema)
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
    const stored = existing && (schema === 1 || baseUpdatedAt !== undefined)
      ? await existing.json()
      : null;

    // A stale schema-1 editor cannot represent 5v5 plays. Keep it compatible
    // with absent/schema-1 playbooks, but never let it downgrade schema 2 and
    // silently replace C-based plays with the old six-player model.
    if (schema === 1 && stored && Number(stored.schema) >= CURRENT_SCHEMA) {
      return jsonNoStore(
        { error: "This editor is out of date. Refresh the page before saving." },
        { status: 422 }
      );
    }

    if (baseUpdatedAt !== undefined) {
      if (existing) {
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
