import { jsonNoStore, requireUser } from "../../_lib/auth.js";
import { finishJobSlot, getJobsBucket, isJobOwner, JOB_ID_RE } from "../../_lib/jobs.js";
import { putJsonIfCurrent } from "../../_lib/r2.js";

const DEFAULT_STALE_MINUTES = 15;

function staleAfterMs(env) {
  const configured = Number.parseInt(env.JOB_STALE_MINUTES, 10);
  const minutes =
    Number.isFinite(configured) && configured >= 10 && configured <= 120
      ? configured
      : DEFAULT_STALE_MINUTES;
  return minutes * 60 * 1000;
}

export async function onRequestGet(context) {
  const { env, params } = context;
  const jobId = Array.isArray(params.jobId) ? params.jobId[0] : params.jobId;

  if (typeof jobId !== "string" || !JOB_ID_RE.test(jobId)) {
    return jsonNoStore({ error: "Invalid job ID" }, { status: 400 });
  }

  try {
    const { user, response } = await requireUser(context);
    if (!user) return response;

    // Return the same not-found response for missing and foreign jobs so this
    // endpoint cannot be used to enumerate another account's job IDs.
    if (!(await isJobOwner(env, jobId, user.userId))) {
      return jsonNoStore({ error: "Job not found" }, { status: 404 });
    }

    const jobsBucket = getJobsBucket(env);
    let object = await jobsBucket.get(`jobs/${jobId}/status.json`);
    if (!object) {
      return jsonNoStore({ error: "Job not found" }, { status: 404 });
    }

    let status = await object.json();
    const createdAtMs = Date.parse(status && status.createdAt);
    if (
      status &&
      status.status === "processing" &&
      Number.isFinite(createdAtMs) &&
      Date.now() - createdAtMs > staleAfterMs(env)
    ) {
      const timedOut = {
        ...status,
        status: "error",
        message: "Processing timed out. Please try generating the playbook again.",
        timedOutAt: new Date().toISOString(),
      };
      const updated = await putJsonIfCurrent(
        { PLAYBOOK_BUCKET: jobsBucket },
        `jobs/${jobId}/status.json`,
        timedOut,
        object
      );
      if (updated !== null) {
        status = timedOut;
      } else {
        // The processor wrote a newer step or completion concurrently. Return
        // that winner rather than falsely reporting a timeout.
        object = await jobsBucket.get(`jobs/${jobId}/status.json`);
        if (!object) return jsonNoStore({ error: "Job not found" }, { status: 404 });
        status = await object.json();
      }
    }

    if (status && (status.status === "complete" || status.status === "error")) {
      const finish = finishJobSlot(env, user.userId, jobId).catch((error) => {
        console.error("Could not release completed job slot:", error);
      });
      if (typeof context.waitUntil === "function") context.waitUntil(finish);
      else await finish;
    }

    return jsonNoStore(status);
  } catch (err) {
    console.error("Status check error:", err);
    return jsonNoStore({ error: "Internal server error" }, { status: 500 });
  }
}
