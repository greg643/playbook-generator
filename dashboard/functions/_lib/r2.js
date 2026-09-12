// Small R2 helpers shared by API handlers. R2 conditional puts are strongly
// consistent and return null when their precondition loses a race.

const JSON_METADATA = { contentType: "application/json" };

export function objectCondition(object) {
  return object
    ? { etagMatches: object.etag }
    : { etagDoesNotMatch: "*" };
}

export async function putJson(env, key, value, options = {}) {
  return env.PLAYBOOK_BUCKET.put(key, JSON.stringify(value), {
    ...options,
    httpMetadata: JSON_METADATA,
  });
}

export async function putJsonIfCurrent(env, key, value, object) {
  return putJson(env, key, value, { onlyIf: objectCondition(object) });
}

export async function createJson(env, key, value) {
  return putJson(env, key, value, {
    onlyIf: { etagDoesNotMatch: "*" },
  });
}

// Inventory first, then delete in batches. Keeping pagination separate from
// deletion avoids cursor surprises if the underlying listing changes while a
// page is being removed, and it also catches unreferenced/orphaned objects.
export async function deleteR2Prefix(bucket, prefix) {
  const keys = [];
  let cursor;
  do {
    const page = await bucket.list({ prefix, limit: 1000, ...(cursor ? { cursor } : {}) });
    for (const object of page.objects || []) keys.push(object.key);
    cursor = page.truncated ? page.cursor : undefined;
    if (page.truncated && !cursor) {
      throw new Error("R2 returned a truncated listing without a cursor");
    }
  } while (cursor);

  for (let i = 0; i < keys.length; i += 1000) {
    await bucket.delete(keys.slice(i, i + 1000));
  }
}
