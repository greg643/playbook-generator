import { isAccountActive, jsonNoStore, requireUser } from "../_lib/auth.js";
import {
  cleanupFailedJob as cleanupFailedJobState,
  createJobOwner,
  getJobsBucket,
  reserveJobSlot,
} from "../_lib/jobs.js";
import { putJson } from "../_lib/r2.js";

const MAX_FILE_BYTES = 50 * 1024 * 1024;
const MAX_REQUEST_BYTES = 52 * 1024 * 1024;
const PPTX_NAME_RE = /^[^/\\\0]{1,128}\.pptx$/i;
const PPTX_CONTENT_TYPE =
  "application/vnd.openxmlformats-officedocument.presentationml.presentation";

function isFile(value) {
  return typeof File !== "undefined" && value instanceof File;
}

function contentLengthTooLarge(request, limit) {
  const raw = request.headers.get("content-length");
  if (!raw) return false;
  return !/^\d+$/.test(raw) || Number(raw) > limit;
}

async function hasZipSignature(file) {
  const signature = new Uint8Array(await file.slice(0, 4).arrayBuffer());
  return (
    signature.length === 4 &&
    signature[0] === 0x50 &&
    signature[1] === 0x4b &&
    signature[2] === 0x03 &&
    signature[3] === 0x04
  );
}

async function cleanupFailedJob(env, userId, jobId, keys) {
  try {
    await cleanupFailedJobState(env, userId, jobId, keys);
  } catch (error) {
    // Keep the quota entry as a durable cleanup index if object scrubbing did
    // not finish; account deletion can then discover and retry this job.
    console.error("Could not completely clean up failed upload job:", error);
  }
}

function quotaResponse(quota) {
  const message =
    quota.reason === "daily"
      ? "Daily generation limit reached; try again tomorrow"
      : "Too many playbooks are already processing; try again shortly";
  return jsonNoStore(
    { error: message },
    { status: 429, headers: { "Retry-After": String(quota.retryAfter) } }
  );
}

export async function onRequestPost(context) {
  const { request, env } = context;

  try {
    const { user, response } = await requireUser(context);
    if (!user) return response;

    const contentType = request.headers.get("content-type") || "";
    if (!contentType.toLowerCase().includes("multipart/form-data")) {
      return jsonNoStore({ error: "Expected multipart/form-data" }, { status: 400 });
    }
    if (contentLengthTooLarge(request, MAX_REQUEST_BYTES)) {
      return jsonNoStore({ error: "Request too large" }, { status: 413 });
    }

    const formData = await request.formData();
    const allowedTextFields = new Set([
      "offense_coach_card",
      "offense_wristband",
      "defense_coach_card",
      "defense_wristband",
    ]);
    const seenTextFields = new Set();
    let textBytes = 0;
    for (const [name, value] of formData.entries()) {
      if (isFile(value)) continue;
      if (
        !allowedTextFields.has(name) ||
        seenTextFields.has(name) ||
        typeof value !== "string" ||
        !["true", "false"].includes(value)
      ) {
        return jsonNoStore({ error: "Unexpected form field" }, { status: 400 });
      }
      seenTextFields.add(name);
      textBytes += new TextEncoder().encode(value).length;
    }
    if (textBytes > 100) {
      return jsonNoStore({ error: "Upload options too large" }, { status: 400 });
    }
    const uploadFields = formData.getAll("file");
    const allFiles = [];
    for (const [, value] of formData.entries()) {
      if (isFile(value)) allFiles.push(value);
    }
    if (uploadFields.length !== 1 || allFiles.length !== 1 || !isFile(uploadFields[0])) {
      return jsonNoStore({ error: "Upload exactly one PowerPoint file" }, { status: 400 });
    }

    const file = uploadFields[0];
    const options = {
      offense_coach_card: formData.get("offense_coach_card") !== "false",
      offense_wristband: formData.get("offense_wristband") !== "false",
      defense_coach_card: formData.get("defense_coach_card") !== "false",
      defense_wristband: formData.get("defense_wristband") !== "false",
    };
    if (!Object.values(options).some(Boolean)) {
      return jsonNoStore({ error: "Select at least one output" }, { status: 400 });
    }

    if (!PPTX_NAME_RE.test(file.name)) {
      return jsonNoStore({ error: "Only .pptx files are accepted" }, { status: 400 });
    }
    if (file.size === 0) {
      return jsonNoStore({ error: "The uploaded PowerPoint file is empty" }, { status: 400 });
    }
    if (file.size > MAX_FILE_BYTES) {
      return jsonNoStore({ error: "File too large (max 50 MB)" }, { status: 413 });
    }
    if (!(await hasZipSignature(file))) {
      return jsonNoStore({ error: "The uploaded file is not a valid PPTX archive" }, { status: 400 });
    }

    const jobId = crypto.randomUUID();
    const quota = await reserveJobSlot(env, user.userId, jobId);
    if (!quota.ok) return quotaResponse(quota);

    const createdAt = new Date().toISOString();
    const jobsBucket = getJobsBucket(env);
    const keys = [
      `jobs/${jobId}/owner.json`,
      `jobs/${jobId}/input.pptx`,
      `jobs/${jobId}/status.json`,
    ];

    try {
      await createJobOwner(env, jobId, user.userId, createdAt);
      // File is a Blob, which R2 accepts directly. Avoid a second 50 MB
      // ArrayBuffer allocation inside the Worker.
      await jobsBucket.put(`jobs/${jobId}/input.pptx`, file, {
        httpMetadata: { contentType: PPTX_CONTENT_TYPE },
      });
      await putJson({ PLAYBOOK_BUCKET: jobsBucket }, `jobs/${jobId}/status.json`, {
        status: "processing",
        createdAt,
        options,
      });

      if (!(await isAccountActive(env, user))) {
        await cleanupFailedJob(env, user.userId, jobId, keys);
        return jsonNoStore({ error: "Account is being deleted" }, { status: 409 });
      }

      const githubResponse = await fetch(
        `https://api.github.com/repos/${env.GITHUB_REPO || "greg643/playbook-generator"}/dispatches`,
        {
          method: "POST",
          headers: {
            Authorization: `Bearer ${env.GITHUB_TOKEN}`,
            Accept: "application/vnd.github.v3+json",
            "User-Agent": "playbook-generator-worker",
          },
          body: JSON.stringify({
            event_type: "process-playbook",
            client_payload: { job_id: jobId },
          }),
        }
      );

      if (!githubResponse.ok) {
        const errorText = await githubResponse.text();
        console.error("GitHub dispatch failed:", githubResponse.status, errorText);
        await cleanupFailedJob(env, user.userId, jobId, keys);
        return jsonNoStore({ error: "Failed to trigger processing" }, { status: 502 });
      }
    } catch (error) {
      await cleanupFailedJob(env, user.userId, jobId, keys);
      throw error;
    }

    return jsonNoStore({ jobId });
  } catch (err) {
    console.error("Upload error:", err);
    return jsonNoStore({ error: "Internal server error" }, { status: 500 });
  }
}
