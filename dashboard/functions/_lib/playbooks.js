import { createJson, deleteR2Prefix } from "./r2.js";

export const DEFAULT_PLAYBOOK_ID = "default";
export const MAX_PLAYBOOKS = 20;
export const MAX_PLAYBOOK_NAME_LENGTH = 60;
export const CATALOG_WRITE_ATTEMPTS = 5;

const UUID_V4_RE = /^[0-9a-f]{8}-[0-9a-f]{4}-4[0-9a-f]{3}-[89ab][0-9a-f]{3}-[0-9a-f]{12}$/;
const CONTROL_CHAR_RE = /[\u0000-\u001f\u007f-\u009f]/u;

export function isPlaybookId(value) {
  return value === DEFAULT_PLAYBOOK_ID ||
    (typeof value === "string" && UUID_V4_RE.test(value));
}

export function normalizePlaybookName(value) {
  if (typeof value !== "string") return null;
  const normalized = value.normalize("NFC").trim();
  if (
    !normalized ||
    Array.from(normalized).length > MAX_PLAYBOOK_NAME_LENGTH ||
    CONTROL_CHAR_RE.test(normalized)
  ) {
    return null;
  }
  return normalized;
}

export function samePlaybookName(left, right) {
  return left.toLocaleLowerCase("en-US") === right.toLocaleLowerCase("en-US");
}

export function legacyPlaybookKey(userId) {
  return `accounts/${userId}/playbook.json`;
}

export function playbookCatalogKey(userId) {
  return `accounts/${userId}/playbooks/catalog.json`;
}

export function playbookItemsPrefix(userId) {
  return `accounts/${userId}/playbooks/items/`;
}

export function playbookObjectKey(userId, playbookId) {
  if (!isPlaybookId(playbookId)) throw new Error("Invalid playbook ID");
  return playbookId === DEFAULT_PLAYBOOK_ID
    ? legacyPlaybookKey(userId)
    : `${playbookItemsPrefix(userId)}${playbookId}.json`;
}

function isTimestamp(value) {
  return typeof value === "string" && value.length <= 40 && Number.isFinite(Date.parse(value));
}

function isCatalogEntry(value) {
  return !!value &&
    typeof value === "object" &&
    !Array.isArray(value) &&
    isPlaybookId(value.id) &&
    normalizePlaybookName(value.name) === value.name &&
    isTimestamp(value.createdAt) &&
    isTimestamp(value.updatedAt) &&
    typeof value.revision === "string" &&
    UUID_V4_RE.test(value.revision);
}

export function isPlaybookCatalog(value) {
  if (
    !value ||
    typeof value !== "object" ||
    Array.isArray(value) ||
    value.schema !== 1 ||
    !isTimestamp(value.updatedAt) ||
    !Array.isArray(value.entries) ||
    value.entries.length < 1 ||
    value.entries.length > MAX_PLAYBOOKS ||
    !value.entries.every(isCatalogEntry)
  ) {
    return false;
  }
  const ids = new Set(value.entries.map((entry) => entry.id));
  const names = new Set(value.entries.map(
    (entry) => entry.name.toLocaleLowerCase("en-US")
  ));
  return ids.size === value.entries.length &&
    names.size === value.entries.length &&
    ids.has(DEFAULT_PLAYBOOK_ID);
}

export function clonePlaybookCatalog(catalog) {
  return {
    schema: 1,
    updatedAt: catalog.updatedAt,
    entries: catalog.entries.map((entry) => ({ ...entry })),
  };
}

function initialCatalog(legacyObject) {
  const now = new Date().toISOString();
  const legacyTime = legacyObject && legacyObject.uploaded instanceof Date
    ? legacyObject.uploaded.toISOString()
    : now;
  return {
    schema: 1,
    updatedAt: now,
    entries: [{
      id: DEFAULT_PLAYBOOK_ID,
      name: "My Playbook",
      createdAt: legacyTime,
      updatedAt: legacyTime,
      revision: crypto.randomUUID(),
    }],
  };
}

export async function readPlaybookCatalog(env, userId) {
  const object = await env.PLAYBOOK_BUCKET.get(playbookCatalogKey(userId));
  if (!object) return { object: null, catalog: null };
  const catalog = await object.json();
  if (!isPlaybookCatalog(catalog)) throw new Error("Invalid playbook catalog");
  return { object, catalog };
}

export async function loadOrCreatePlaybookCatalog(env, userId) {
  for (let attempt = 0; attempt < CATALOG_WRITE_ATTEMPTS; attempt += 1) {
    const loaded = await readPlaybookCatalog(env, userId);
    if (loaded.catalog) return loaded;

    const legacyObject = await env.PLAYBOOK_BUCKET.head(legacyPlaybookKey(userId));
    const catalog = initialCatalog(legacyObject);
    const object = await createJson(env, playbookCatalogKey(userId), catalog);
    if (object !== null) return { object, catalog };
  }
  throw new Error("Could not initialize playbook catalog after concurrent updates");
}

export function findPlaybookEntry(catalog, playbookId) {
  return catalog.entries.find((entry) => entry.id === playbookId) || null;
}

export async function deleteAccountPlaybooks(env, userId) {
  await env.PLAYBOOK_BUCKET.delete(legacyPlaybookKey(userId));
  await deleteR2Prefix(env.PLAYBOOK_BUCKET, `accounts/${userId}/playbooks/`);
}
