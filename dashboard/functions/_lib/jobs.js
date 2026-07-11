import { createJson, putJson, putJsonIfCurrent } from "./r2.js";

export const JOB_ID_RE = /^[0-9a-f]{8}-[0-9a-f]{4}-[1-8][0-9a-f]{3}-[89ab][0-9a-f]{3}-[0-9a-f]{12}$/i;
export const PDF_NAME_RE = /^[A-Za-z0-9][A-Za-z0-9._-]{0,126}\.pdf$/;

const DEFAULT_MAX_ACTIVE_JOBS = 2;
const DEFAULT_MAX_DAILY_JOBS = 20;
const DEFAULT_ACTIVE_MINUTES = 30;
const MAX_WRITE_ATTEMPTS = 5;

// Deployments can bind a job-only bucket as JOBS_BUCKET so the untrusted file
// processing runner never needs credentials for account/authentication data.
// Fall back to the original binding for a no-downtime migration.
export function getJobsBucket(env) {
  return env.JOBS_BUCKET || env.PLAYBOOK_BUCKET;
}

function boundedEnvInt(value, fallback, min, max) {
  const parsed = Number.parseInt(value, 10);
  return Number.isFinite(parsed) && parsed >= min && parsed <= max ? parsed : fallback;
}

function quotaSettings(env) {
  return {
    maxActive: boundedEnvInt(env.MAX_ACTIVE_JOBS_PER_USER, DEFAULT_MAX_ACTIVE_JOBS, 1, 10),
    maxDaily: boundedEnvInt(env.MAX_DAILY_JOBS_PER_USER, DEFAULT_MAX_DAILY_JOBS, 1, 200),
    activeMs:
      boundedEnvInt(env.JOB_ACTIVE_MINUTES, DEFAULT_ACTIVE_MINUTES, 5, 120) * 60 * 1000,
  };
}

function dayString(nowMs) {
  return new Date(nowMs).toISOString().slice(0, 10);
}

function quotaKey(userId, nowMs) {
  return `accounts/${userId}/job-quota/${dayString(nowMs)}.json`;
}

async function readQuota(env, key, date) {
  const object = await env.PLAYBOOK_BUCKET.get(key);
  if (!object) return { object: null, record: { schema: 1, date, jobs: [] } };

  let record;
  try {
    record = await object.json();
  } catch (error) {
    throw new Error("Invalid job quota record", { cause: error });
  }
  if (!record || record.date !== date || !Array.isArray(record.jobs)) {
    throw new Error("Invalid job quota record");
  }
  return { object, record };
}

function validQuotaJob(job) {
  return (
    job &&
    typeof job === "object" &&
    JOB_ID_RE.test(job.id) &&
    typeof job.createdAt === "string" &&
    typeof job.activeUntil === "string"
  );
}

export async function reserveJobSlot(env, userId, jobId, nowMs = Date.now()) {
  const settings = quotaSettings(env);
  const date = dayString(nowMs);
  const key = quotaKey(userId, nowMs);

  for (let attempt = 0; attempt < MAX_WRITE_ATTEMPTS; attempt += 1) {
    const { object, record } = await readQuota(env, key, date);
    const jobs = record.jobs.filter(validQuotaJob);
    const active = jobs.filter(
      (job) => !job.finishedAt && Date.parse(job.activeUntil) > nowMs
    );

    if (jobs.length >= settings.maxDaily) {
      const nextDayMs = Date.parse(`${date}T00:00:00.000Z`) + 24 * 60 * 60 * 1000;
      return {
        ok: false,
        reason: "daily",
        retryAfter: Math.max(1, Math.ceil((nextDayMs - nowMs) / 1000)),
      };
    }
    if (active.length >= settings.maxActive) {
      const nextMs = Math.min(...active.map((job) => Date.parse(job.activeUntil)));
      return {
        ok: false,
        reason: "active",
        retryAfter: Math.max(1, Math.ceil((nextMs - nowMs) / 1000)),
      };
    }

    const createdAt = new Date(nowMs).toISOString();
    jobs.push({
      id: jobId,
      createdAt,
      activeUntil: new Date(nowMs + settings.activeMs).toISOString(),
    });
    const result = await putJsonIfCurrent(env, key, { schema: 1, date, jobs }, object);
    if (result !== null) return { ok: true };
  }

  throw new Error("Could not reserve a job slot after concurrent updates");
}

async function updateReservation(env, userId, jobId, { remove = false, nowMs = Date.now() } = {}) {
  const date = dayString(nowMs);
  const key = quotaKey(userId, nowMs);

  for (let attempt = 0; attempt < MAX_WRITE_ATTEMPTS; attempt += 1) {
    const { object, record } = await readQuota(env, key, date);
    if (!object) return;

    const index = record.jobs.findIndex((job) => job && job.id === jobId);
    if (index === -1) return;
    const jobs = record.jobs.slice();
    if (remove) {
      jobs.splice(index, 1);
    } else {
      jobs[index] = {
        ...jobs[index],
        activeUntil: new Date(nowMs).toISOString(),
        finishedAt: new Date(nowMs).toISOString(),
      };
    }

    const result = await putJsonIfCurrent(env, key, { ...record, jobs }, object);
    if (result !== null) return;
  }
  throw new Error("Could not update a job slot after concurrent updates");
}

export function finishJobSlot(env, userId, jobId, nowMs) {
  return updateReservation(env, userId, jobId, { nowMs });
}

export function cancelJobSlot(env, userId, jobId, nowMs) {
  return updateReservation(env, userId, jobId, { remove: true, nowMs });
}

export async function listUserJobIds(env, userId, nowMs = Date.now()) {
  const jobIds = new Set();
  // The job bucket has a mandatory one-day lifecycle. Read a bounded two-day
  // window to cover UTC midnight and delayed cleanup without making account
  // deletion scale with account age. Historical quota keys are removed below
  // without loading their bodies.
  const dates = [dayString(nowMs), dayString(nowMs - 24 * 60 * 60 * 1000)];
  const records = await Promise.all(
    dates.map(async (date) => {
      try {
        const object = await env.PLAYBOOK_BUCKET.get(
          `accounts/${userId}/job-quota/${date}.json`
        );
        return object ? await object.json() : null;
      } catch (error) {
        throw new Error("Could not read quota record during account cleanup", {
          cause: error,
        });
      }
    })
  );
  for (const record of records) {
    if (!record || !Array.isArray(record.jobs)) continue;
    for (const job of record.jobs) {
      if (job && typeof job.id === "string" && JOB_ID_RE.test(job.id)) {
        jobIds.add(job.id);
      }
    }
  }
  return [...jobIds];
}

export async function deleteUserQuotaRecords(env, userId, knownJobIds = null) {
  const prefix = `accounts/${userId}/job-quota/`;
  const jobIds = knownJobIds ? [...new Set(knownJobIds)] : await listUserJobIds(env, userId);
  // Re-list from the start after each batch. This remains correct while the
  // preceding objects are being deleted and avoids relying on cursors whose
  // earlier keys no longer exist.
  while (true) {
    const page = await env.PLAYBOOK_BUCKET.list({ prefix, limit: 1000 });
    const keys = (page.objects || []).map((object) => object.key);
    if (keys.length === 0) return jobIds;
    await env.PLAYBOOK_BUCKET.delete(keys);
  }
}

async function scrubJobPayload(bucket, jobId) {
  const prefix = `jobs/${jobId}/`;
  const preserved = new Set([
    `${prefix}owner.json`,
    `${prefix}cancelled.json`,
    `${prefix}status.json`,
  ]);
  while (true) {
    const page = await bucket.list({ prefix, limit: 1000 });
    const keys = (page.objects || [])
      .map((object) => object.key)
      .filter((key) => !preserved.has(key));
    if (keys.length === 0) return;
    await bucket.delete(keys);
  }
}

// Leave a small cancellation tombstone so an already-dispatched worker cannot
// recreate deleted uploads/PDFs. Job-bucket lifecycle rules remove these
// non-content metadata objects after the configured retention window.
export async function cancelAndScrubUserJobs(
  env,
  userId,
  jobIds,
  now = new Date(),
  reason = "account-deleted"
) {
  const bucket = getJobsBucket(env);
  const cancelledAt = now.toISOString();
  const scrubbed = [];
  for (const jobId of new Set(jobIds || [])) {
    if (!JOB_ID_RE.test(jobId)) continue;
    const ownerObject = await bucket.get(`jobs/${jobId}/owner.json`);
    if (!ownerObject) {
      const cancellationObject = await bucket.get(`jobs/${jobId}/cancelled.json`);
      if (cancellationObject) {
        const cancellation = await cancellationObject.json();
        if (!cancellation || cancellation.ownerId !== userId) {
          throw new Error("Job cancellation ownership changed during account cleanup");
        }
        await scrubJobPayload(bucket, jobId);
        await putJson({ PLAYBOOK_BUCKET: bucket }, `jobs/${jobId}/status.json`, {
          status: "error",
          message:
            reason === "account-deleted"
              ? "This job was cancelled because its account was deleted."
              : "This job was cancelled before processing could start.",
          cancelledAt: cancellation.cancelledAt || cancelledAt,
        });
        scrubbed.push(jobId);
        continue;
      }
      const page = await bucket.list({ prefix: `jobs/${jobId}/`, limit: 1 });
      if ((page.objects || []).length === 0) {
        // The quota reservation can become visible just before its owner write.
        // Place the cancellation marker now so that in-flight setup or a
        // dispatched worker cannot populate this prefix after account cleanup.
        await putJson({ PLAYBOOK_BUCKET: bucket }, `jobs/${jobId}/cancelled.json`, {
          schema: 1,
          ownerId: userId,
          cancelledAt,
          reason,
        });
        await scrubJobPayload(bucket, jobId);
        await putJson({ PLAYBOOK_BUCKET: bucket }, `jobs/${jobId}/status.json`, {
          status: "error",
          message:
            reason === "account-deleted"
              ? "This job was cancelled because its account was deleted."
              : "This job was cancelled before processing could start.",
          cancelledAt,
        });
        scrubbed.push(jobId);
        continue;
      }
      throw new Error("Could not verify job ownership during account cleanup");
    }
    let owner;
    try {
      owner = await ownerObject.json();
    } catch (error) {
      throw new Error("Invalid job owner during account cleanup", { cause: error });
    }
    if (!owner || owner.ownerId !== userId) {
      throw new Error("Job ownership changed during account cleanup");
    }

    await putJson({ PLAYBOOK_BUCKET: bucket }, `jobs/${jobId}/cancelled.json`, {
      schema: 1,
      ownerId: userId,
      cancelledAt,
      reason,
    });
    await scrubJobPayload(bucket, jobId);
    await putJson({ PLAYBOOK_BUCKET: bucket }, `jobs/${jobId}/status.json`, {
      status: "error",
      message:
        reason === "account-deleted"
          ? "This job was cancelled because its account was deleted."
          : "This job was cancelled before processing could start.",
      cancelledAt,
    });
    scrubbed.push(jobId);
  }
  return scrubbed;
}

// Failed setup/dispatch cleanup keeps the quota record as a durable account→job
// index until payload scrubbing has definitely succeeded. Account deletion can
// therefore rediscover and retry any cleanup interrupted by an R2 error.
export async function cleanupFailedJob(env, userId, jobId, keys) {
  const bucket = getJobsBucket(env);
  const owner = await bucket.get(`jobs/${jobId}/owner.json`);
  if (owner) {
    const scrubbed = await cancelAndScrubUserJobs(
      env,
      userId,
      [jobId],
      new Date(),
      "setup-failed"
    );
    if (!scrubbed.includes(jobId)) throw new Error("Failed job was not scrubbed");
  } else {
    // Owner creation is the first job-bucket write. If it did not complete,
    // remove/verify the finite set of keys this request could have attempted.
    if (keys.length) await bucket.delete(keys);
    const remaining = await Promise.all(keys.map((key) => bucket.get(key)));
    if (remaining.some(Boolean)) throw new Error("Failed job payload still exists");
  }

  await cancelJobSlot(env, userId, jobId);
}

export async function createJobOwner(env, jobId, userId, createdAt) {
  const jobsEnv = { PLAYBOOK_BUCKET: getJobsBucket(env) };
  const result = await createJson(jobsEnv, `jobs/${jobId}/owner.json`, {
    schema: 1,
    ownerId: userId,
    createdAt,
  });
  if (result === null) throw new Error("Job ID collision");
}

export async function isJobOwner(env, jobId, userId) {
  const object = await getJobsBucket(env).get(`jobs/${jobId}/owner.json`);
  if (!object) return false;
  try {
    const owner = await object.json();
    return owner && owner.ownerId === userId;
  } catch (error) {
    console.error("Invalid job owner metadata:", error);
    return false;
  }
}
