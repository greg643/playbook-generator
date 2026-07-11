import {
  jsonNoStore,
  requireUser,
  emailKey,
  isValidEmail,
  normalizeEmail,
  hashPassword,
  constantTimeEqualHex,
  requestBodyTooLarge,
  utf8ByteLength,
} from "../../_lib/auth.js";
import { createJson, putJsonIfCurrent } from "../../_lib/r2.js";
import {
  cancelAndScrubUserJobs,
  deleteUserQuotaRecords,
  JOB_ID_RE,
  listUserJobIds,
} from "../../_lib/jobs.js";

const MAX_BODY_BYTES = 10000;
const MAX_INDEX_ATTEMPTS = 5;

function deletionKey(userId) {
  return `accounts/${userId}/deletion.json`;
}

function validDeletionRecord(value, userId) {
  return (
    value &&
    value.schema === 1 &&
    value.userId === userId &&
    Array.isArray(value.jobIds) &&
    value.jobIds.every((jobId) => typeof jobId === "string" && JOB_ID_RE.test(jobId))
  );
}

async function ensureDeletionRecord(env, userId) {
  const key = deletionKey(userId);
  for (let attempt = 0; attempt < MAX_INDEX_ATTEMPTS; attempt += 1) {
    const object = await env.PLAYBOOK_BUCKET.get(key);
    let existing = null;
    if (object) {
      existing = await object.json();
      if (!validDeletionRecord(existing, userId)) {
        throw new Error("Invalid account deletion record");
      }
    }

    // Quota records are the durable user→job index. Copy their IDs into the
    // deletion record before removing any data so retries remain possible even
    // after quota cleanup has completed.
    const listedJobIds = await listUserJobIds(env, userId);
    const jobIds = [...new Set([...(existing ? existing.jobIds : []), ...listedJobIds])];
    const record = {
      schema: 1,
      userId,
      jobIds,
      startedAt: (existing && existing.startedAt) || new Date().toISOString(),
    };
    if (existing && existing.jobIds.length === jobIds.length) return record;

    const written = object
      ? await putJsonIfCurrent(env, key, record, object)
      : await createJson(env, key, record);
    if (written !== null) return record;
  }
  throw new Error("Could not create account deletion record after concurrent updates");
}

async function completeDeletion(env, credentialKey, record) {
  // Close the inventory race: a job that reserved before disabledAt but after
  // the first scan is now visible. Reservations made later observe disabledAt
  // and self-clean instead of dispatching.
  const deletion = await ensureDeletionRecord(env, record.userId);
  await env.PLAYBOOK_BUCKET.delete(`accounts/${record.userId}/playbook.json`);
  await cancelAndScrubUserJobs(env, record.userId, deletion.jobIds);
  await deleteUserQuotaRecords(env, record.userId, deletion.jobIds);
  await env.PLAYBOOK_BUCKET.delete(`accounts/${record.userId}/playbook.json`);
  await env.PLAYBOOK_BUCKET.delete(deletionKey(record.userId));

  // A minimal conditional tombstone allows email reuse without letting a
  // delayed deletion erase the replacement account.
  const credentialObject = await env.PLAYBOOK_BUCKET.get(credentialKey);
  if (credentialObject) {
    const current = await credentialObject.json();
    if (current && current.userId === record.userId) {
      const finalized = await putJsonIfCurrent(
        env,
        credentialKey,
        { schema: 1, deletedAt: new Date().toISOString() },
        credentialObject
      );
      if (finalized === null) throw new Error("Account finalization lost a concurrent update");
    }
  }
  return deletion.jobIds;
}

function delay(ms) {
  return new Promise((resolve) => setTimeout(resolve, ms));
}

async function reconcileDeletion(env, credentialKey, record) {
  let lastError;
  for (const waitMs of [0, 1000, 3000]) {
    if (waitMs) await delay(waitMs);
    try {
      const jobIds = await completeDeletion(env, credentialKey, record);
      // One delayed sweep catches a mutation that authenticated immediately
      // before disabledAt and finished after the foreground cleanup.
      await delay(1000);
      await env.PLAYBOOK_BUCKET.delete(`accounts/${record.userId}/playbook.json`);
      await cancelAndScrubUserJobs(env, record.userId, jobIds);
      await deleteUserQuotaRecords(env, record.userId, jobIds);
      return;
    } catch (error) {
      lastError = error;
    }
  }
  throw lastError;
}

export async function onRequestPost(context) {
  const { request, env } = context;

  try {
    if (requestBodyTooLarge(request, MAX_BODY_BYTES)) {
      return jsonNoStore({ error: "Request too large" }, { status: 413 });
    }
    const bodyText = await request.text();
    if (utf8ByteLength(bodyText) > MAX_BODY_BYTES) {
      return jsonNoStore({ error: "Request too large" }, { status: 413 });
    }

    let body;
    try {
      body = JSON.parse(bodyText);
    } catch (err) {
      return jsonNoStore({ error: "Invalid JSON" }, { status: 400 });
    }

    const password = body.password;
    const requestedUserId = body.userId;
    if (
      typeof password !== "string" ||
      password.length > 1024 ||
      typeof requestedUserId !== "string"
    ) {
      return jsonNoStore({ error: "Invalid password" }, { status: 401 });
    }

    // A deletion tombstone blocks every normal endpoint, but the original
    // same-version session may re-enter this endpoint. If that cookie is gone,
    // email + password may resume only an already-disabled deletion; it cannot
    // initiate deletion for an active account.
    let { user, response } = await requireUser(context, { allowDisabled: true });
    if (!user) {
      const email = normalizeEmail(body.email);
      if (!isValidEmail(email)) return response;
      const accountObject = await env.PLAYBOOK_BUCKET.get(await emailKey(email));
      if (!accountObject) return response;
      const account = await accountObject.json();
      if (
        !account ||
        !account.disabledAt ||
        typeof account.userId !== "string" ||
        account.userId !== requestedUserId ||
        normalizeEmail(account.email) !== email
      ) {
        return response;
      }
      user = { userId: account.userId, email, account, accountObject };
    }
    if (user.userId !== requestedUserId) {
      return jsonNoStore({ error: "Account changed while deletion was pending" }, { status: 409 });
    }

    const key = await emailKey(user.email);
    const record = { ...user.account };
    const hash = await hashPassword(password, record.salt, record.iterations);
    if (!constantTimeEqualHex(hash, record.hash)) {
      return jsonNoStore({ error: "Invalid password" }, { status: 401 });
    }

    await ensureDeletionRecord(env, record.userId);

    if (!record.disabledAt) {
      // Keep the session version unchanged: disabledAt rejects this cookie from
      // every normal endpoint, while delete-account can use it for idempotent
      // retries if storage cleanup is temporarily unavailable.
      record.disabledAt = new Date().toISOString();
      record.deletionPendingAt = record.disabledAt;
      const disabled = await putJsonIfCurrent(env, key, record, user.accountObject);
      if (disabled === null) {
        return jsonNoStore(
          { error: "Account changed while it was being deleted; please sign in and try again" },
          { status: 409 }
        );
      }
    }

    try {
      const jobIds = await completeDeletion(env, key, record);
      if (typeof context.waitUntil === "function") {
        context.waitUntil(
          (async () => {
            await delay(1000);
            await env.PLAYBOOK_BUCKET.delete(`accounts/${record.userId}/playbook.json`);
            await cancelAndScrubUserJobs(env, record.userId, jobIds);
            await deleteUserQuotaRecords(env, record.userId, jobIds);
          })().catch((error) => console.error("Delayed account resweep failed:", error))
        );
      }
    } catch (cleanupError) {
      console.error("Account cleanup is still pending:", cleanupError);
      const latest = await env.PLAYBOOK_BUCKET.get(key);
      if (!latest) {
        // Deleting the credential may have succeeded even if the storage call
        // reported a transport error. In that case the deletion is complete.
        return jsonNoStore({ ok: true });
      }
      try {
        const latestRecord = await latest.json();
        if (latestRecord.deletedAt || latestRecord.userId !== record.userId) {
          return jsonNoStore({ ok: true });
        }
      } catch (readError) {
        console.error("Could not inspect pending account deletion:", readError);
      }
      if (typeof context.waitUntil === "function") {
        context.waitUntil(
          reconcileDeletion(env, key, record).catch((error) =>
            console.error("Background account reconciliation failed:", error)
          )
        );
      }
      return jsonNoStore(
        {
          error:
            "Account access is disabled, but cleanup could not finish. Retry deletion to complete it.",
          deletionPending: true,
        },
        { status: 503, headers: { "Retry-After": "5" } }
      );
    }

    return jsonNoStore({ ok: true });
  } catch (err) {
    console.error("Delete-account error:", err);
    return jsonNoStore({ error: "Internal server error" }, { status: 500 });
  }
}
