import {
  isAccountActive,
  jsonNoStore,
  requestBodyTooLarge,
  requireUser,
  utf8ByteLength,
} from "../_lib/auth.js";
import {
  CATALOG_WRITE_ATTEMPTS,
  clonePlaybookCatalog,
  deleteAccountPlaybooks,
  findPlaybookEntry,
  isPlaybookId,
  loadOrCreatePlaybookCatalog,
  MAX_PLAYBOOKS,
  normalizePlaybookName,
  playbookCatalogKey,
  playbookObjectKey,
  readPlaybookCatalog,
  samePlaybookName,
} from "../_lib/playbooks.js";
import { createJson, putJsonIfCurrent } from "../_lib/r2.js";

const MAX_BODY_BYTES = 10000;
const SUPPORTED_PLAYER_COUNTS = new Set([5, 6]);

async function readBody(request) {
  if (requestBodyTooLarge(request, MAX_BODY_BYTES)) {
    return { response: jsonNoStore({ error: "Request too large" }, { status: 413 }) };
  }
  const text = await request.text();
  if (utf8ByteLength(text) > MAX_BODY_BYTES) {
    return { response: jsonNoStore({ error: "Request too large" }, { status: 413 }) };
  }
  try {
    return { body: JSON.parse(text) };
  } catch {
    return { response: jsonNoStore({ error: "Invalid JSON" }, { status: 400 }) };
  }
}

function publicCatalog(catalog) {
  return {
    playbooks: catalog.entries,
    maxPlaybooks: MAX_PLAYBOOKS,
  };
}

async function finishResponse(env, user, response) {
  if (!(await isAccountActive(env, user))) {
    try {
      await deleteAccountPlaybooks(env, user.userId);
    } catch (error) {
      // Account deletion owns the durable retry path. Never return a normal
      // catalog response once the credential has been disabled, even if this
      // request's best-effort resweep encounters a transient storage failure.
      console.error("Could not resweep disabled account playbooks:", error);
    }
    return jsonNoStore({ error: "Account is being deleted" }, { status: 409 });
  }
  return response;
}

async function finishMutation(env, user, catalog, options = {}) {
  return finishResponse(env, user, jsonNoStore(publicCatalog(catalog), options));
}

export async function onRequestGet(context) {
  const { env } = context;
  try {
    const { user, response } = await requireUser(context);
    if (!user) return response;
    const { catalog } = await loadOrCreatePlaybookCatalog(env, user.userId);
    return finishMutation(env, user, catalog);
  } catch (error) {
    console.error("Playbooks GET error:", error);
    return jsonNoStore({ error: "Internal server error" }, { status: 500 });
  }
}

export async function onRequestPost(context) {
  const { request, env } = context;
  let user = null;
  let createdKey = null;
  let createdUserId = null;
  let createdPlaybookId = null;
  let published = false;
  try {
    const auth = await requireUser(context);
    if (!auth.user) return auth.response;
    user = auth.user;
    const parsed = await readBody(request);
    if (parsed.response) return parsed.response;
    const name = normalizePlaybookName(parsed.body && parsed.body.name);
    const defaultPlayersPerSide = parsed.body && parsed.body.defaultPlayersPerSide;
    if (!name || !SUPPORTED_PLAYER_COUNTS.has(defaultPlayersPerSide)) {
      return jsonNoStore({ error: "Enter a playbook name and choose 5v5 or 6v6" }, { status: 400 });
    }

    const initial = await loadOrCreatePlaybookCatalog(env, user.userId);
    if (initial.catalog.entries.length >= MAX_PLAYBOOKS) {
      return finishResponse(
        env,
        user,
        jsonNoStore({ error: `Playbook limit reached (max ${MAX_PLAYBOOKS})` }, { status: 409 })
      );
    }
    if (initial.catalog.entries.some((entry) => samePlaybookName(entry.name, name))) {
      return finishResponse(
        env,
        user,
        jsonNoStore({ error: "A playbook with that name already exists" }, { status: 409 })
      );
    }

    const now = new Date().toISOString();
    const playbook = {
      id: crypto.randomUUID(),
      name,
      createdAt: now,
      updatedAt: now,
      revision: crypto.randomUUID(),
    };
    const document = {
      schema: 2,
      defaultPlayersPerSide,
      offense: [],
      defense: [],
      updatedAt: now,
      revision: crypto.randomUUID(),
    };
    createdKey = playbookObjectKey(user.userId, playbook.id);
    createdUserId = user.userId;
    createdPlaybookId = playbook.id;
    if (await createJson(env, createdKey, document) === null) {
      throw new Error("Generated playbook ID already exists");
    }

    for (let attempt = 0; attempt < CATALOG_WRITE_ATTEMPTS; attempt += 1) {
      const loaded = await loadOrCreatePlaybookCatalog(env, user.userId);
      if (findPlaybookEntry(loaded.catalog, playbook.id)) {
        published = true;
        return finishMutation(env, user, loaded.catalog, { status: 201 });
      }
      if (loaded.catalog.entries.length >= MAX_PLAYBOOKS) {
        await env.PLAYBOOK_BUCKET.delete(createdKey);
        createdKey = null;
        return finishResponse(
          env,
          user,
          jsonNoStore({ error: `Playbook limit reached (max ${MAX_PLAYBOOKS})` }, { status: 409 })
        );
      }
      if (loaded.catalog.entries.some((entry) => samePlaybookName(entry.name, name))) {
        await env.PLAYBOOK_BUCKET.delete(createdKey);
        createdKey = null;
        return finishResponse(
          env,
          user,
          jsonNoStore({ error: "A playbook with that name already exists" }, { status: 409 })
        );
      }
      const next = clonePlaybookCatalog(loaded.catalog);
      next.entries.push(playbook);
      next.updatedAt = new Date().toISOString();
      const saved = await putJsonIfCurrent(
        env,
        playbookCatalogKey(user.userId),
        next,
        loaded.object
      );
      if (saved !== null) {
        published = true;
        return finishMutation(env, user, next, { status: 201 });
      }
    }

    await env.PLAYBOOK_BUCKET.delete(createdKey);
    createdKey = null;
    return finishResponse(
      env,
      user,
      jsonNoStore(
        { error: "Playbook list changed repeatedly; please try again" },
        { status: 503, headers: { "Retry-After": "1" } }
      )
    );
  } catch (error) {
    // If a transport error happened after the catalog write, confirm whether
    // the entry is visible before treating its document as an orphan.
    if (createdKey && !published) {
      try {
        const loaded = await readPlaybookCatalog(env, createdUserId);
        if (
          loaded.catalog &&
          findPlaybookEntry(loaded.catalog, createdPlaybookId)
        ) {
          published = true;
          return finishMutation(env, user, loaded.catalog, { status: 201 });
        }
        if (!published) await env.PLAYBOOK_BUCKET.delete(createdKey);
      } catch (cleanupError) {
        console.error("Could not reconcile failed playbook creation:", cleanupError);
      }
    }
    console.error("Playbooks POST error:", error);
    const failure = jsonNoStore({ error: "Internal server error" }, { status: 500 });
    return user ? finishResponse(env, user, failure) : failure;
  }
}

export async function onRequestPatch(context) {
  const { request, env } = context;
  let user = null;
  try {
    const auth = await requireUser(context);
    if (!auth.user) return auth.response;
    user = auth.user;
    const parsed = await readBody(request);
    if (parsed.response) return parsed.response;
    const playbookId = parsed.body && parsed.body.playbookId;
    const name = normalizePlaybookName(parsed.body && parsed.body.name);
    const baseRevision = parsed.body && parsed.body.baseRevision;
    if (!isPlaybookId(playbookId) || !name || typeof baseRevision !== "string") {
      return jsonNoStore({ error: "Invalid playbook update" }, { status: 400 });
    }

    for (let attempt = 0; attempt < CATALOG_WRITE_ATTEMPTS; attempt += 1) {
      const loaded = await loadOrCreatePlaybookCatalog(env, user.userId);
      const current = findPlaybookEntry(loaded.catalog, playbookId);
      if (!current) {
        return finishResponse(
          env,
          user,
          jsonNoStore({ error: "Playbook not found" }, { status: 404 })
        );
      }
      if (current.revision !== baseRevision) {
        return finishResponse(
          env,
          user,
          jsonNoStore(
            { error: "conflict", playbook: current },
            { status: 409 }
          )
        );
      }
      if (loaded.catalog.entries.some(
        (entry) => entry.id !== playbookId && samePlaybookName(entry.name, name)
      )) {
        return finishResponse(
          env,
          user,
          jsonNoStore({ error: "A playbook with that name already exists" }, { status: 409 })
        );
      }
      if (current.name === name) return finishMutation(env, user, loaded.catalog);

      const now = new Date().toISOString();
      const renamed = { ...current, name, updatedAt: now, revision: crypto.randomUUID() };
      const next = clonePlaybookCatalog(loaded.catalog);
      next.entries = next.entries.map((entry) => entry.id === playbookId ? renamed : entry);
      next.updatedAt = now;
      const saved = await putJsonIfCurrent(
        env,
        playbookCatalogKey(user.userId),
        next,
        loaded.object
      );
      if (saved !== null) return finishMutation(env, user, next);
    }
    return finishResponse(
      env,
      user,
      jsonNoStore(
        { error: "Playbook list changed repeatedly; please try again" },
        { status: 503, headers: { "Retry-After": "1" } }
      )
    );
  } catch (error) {
    console.error("Playbooks PATCH error:", error);
    const failure = jsonNoStore({ error: "Internal server error" }, { status: 500 });
    return user ? finishResponse(env, user, failure) : failure;
  }
}
