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
