import { isAccountActive, jsonNoStore, requireUser } from "../_lib/auth.js";
import {
  cleanupFailedJob as cleanupFailedJobState,
  createJobOwner,
  getJobsBucket,
  reserveJobSlot,
} from "../_lib/jobs.js";
import { putJson } from "../_lib/r2.js";

const OFFENSE_NAME_RE = /^(0[1-9]|1[0-6])\.png$/;
const DEFENSE_NAME_RE = /^D[1-6]\.png$/;
const MAX_FILES = 22;
const MAX_FILE_BYTES = 4 * 1024 * 1024;
const MAX_TOTAL_BYTES = 60 * 1024 * 1024;
const MAX_REQUEST_BYTES = 62 * 1024 * 1024;

function isFile(value) {
  return typeof File !== "undefined" && value instanceof File;
}

function contentLengthTooLarge(request, limit) {
  const raw = request.headers.get("content-length");
  if (!raw) return false;
  return !/^\d+$/.test(raw) || Number(raw) > limit;
}

async function hasPngSignature(file) {
  const bytes = new Uint8Array(await file.slice(0, 8).arrayBuffer());
  const signature = [0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a];
  return bytes.length === signature.length && signature.every((byte, i) => bytes[i] === byte);
}

async function cleanupFailedJob(env, userId, jobId, keys) {
  try {
    await cleanupFailedJobState(env, userId, jobId, keys);
  } catch (error) {
    // The quota entry intentionally remains as a durable cleanup index. The
    // job-bucket lifecycle is the final backstop if R2 stays unavailable.
    console.error("Could not completely clean up failed generation job:", error);
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

    let optionsFieldCount = 0;
    let textBytes = 0;
    for (const [name, value] of formData.entries()) {
      if (isFile(value)) continue;
      if (name !== "options" || typeof value !== "string") {
        return jsonNoStore({ error: "Unexpected form field" }, { status: 400 });
      }
      optionsFieldCount += 1;
      textBytes += new TextEncoder().encode(value).length;
    }
    if (optionsFieldCount !== 1 || textBytes > 1000) {
      return jsonNoStore({ error: "Invalid options" }, { status: 400 });
    }

    let options;
    try {
      options = JSON.parse(formData.get("options"));
    } catch (err) {
      return jsonNoStore({ error: "Invalid options" }, { status: 400 });
    }
    if (!options || typeof options !== "object" || Array.isArray(options)) {
      return jsonNoStore({ error: "Invalid options" }, { status: 400 });
    }
    const selected = {
      offense_coach_card: options.offense_coach_card === true,
      offense_wristband: options.offense_wristband === true,
      defense_coach_card: options.defense_coach_card === true,
      defense_wristband: options.defense_wristband === true,
    };
    if (!Object.values(selected).some(Boolean)) {
      return jsonNoStore({ error: "Select at least one output" }, { status: 400 });
    }
    // Title flags ride along after the output check so they can't satisfy it.
    selected.show_offense_title = options.show_offense_title === true;
    selected.show_defense_title = options.show_defense_title !== false;

    const files = formData.getAll("plays");
    const allFiles = [];
    for (const [, value] of formData.entries()) {
      if (isFile(value)) allFiles.push(value);
    }
    if (files.length === 0) {
      return jsonNoStore({ error: "No play images uploaded" }, { status: 400 });
    }
    if (files.length > MAX_FILES || allFiles.length !== files.length) {
      return jsonNoStore(
        { error: `Upload at most ${MAX_FILES} play images and no extra files` },
        { status: 400 }
      );
    }

    let totalBytes = 0;
    let hasOffense = false;
    let hasDefense = false;
    const names = new Set();
    for (const file of files) {
      if (!isFile(file)) {
        return jsonNoStore({ error: "Invalid play upload" }, { status: 400 });
      }
      if (names.has(file.name)) {
        return jsonNoStore({ error: `Duplicate play file name: ${file.name}` }, { status: 400 });
      }
      names.add(file.name);

      if (OFFENSE_NAME_RE.test(file.name)) hasOffense = true;
      else if (DEFENSE_NAME_RE.test(file.name)) hasDefense = true;
      else {
        return jsonNoStore({ error: `Invalid play file name: ${file.name}` }, { status: 400 });
      }
      if (file.type !== "image/png") {
        return jsonNoStore({ error: "Play images must be PNG" }, { status: 400 });
      }
      if (file.size === 0) {
        return jsonNoStore({ error: `Play image is empty: ${file.name}` }, { status: 400 });
      }
      if (file.size > MAX_FILE_BYTES) {
        return jsonNoStore(
          { error: `Play image too large (max 4 MB): ${file.name}` },
          { status: 413 }
        );
      }
      if (!(await hasPngSignature(file))) {
        return jsonNoStore({ error: `Invalid PNG data: ${file.name}` }, { status: 400 });
      }
      totalBytes += file.size;
    }
    if (totalBytes > MAX_TOTAL_BYTES) {
      return jsonNoStore({ error: "Total upload too large (max 60 MB)" }, { status: 413 });
    }
    if ((selected.offense_coach_card || selected.offense_wristband) && !hasOffense) {
      return jsonNoStore(
        { error: "Offense outputs selected but no offense plays" },
        { status: 400 }
      );
    }
    if ((selected.defense_coach_card || selected.defense_wristband) && !hasDefense) {
      return jsonNoStore(
        { error: "Defense outputs selected but no defense plays" },
        { status: 400 }
      );
    }

    const jobId = crypto.randomUUID();
    const quota = await reserveJobSlot(env, user.userId, jobId);
    if (!quota.ok) return quotaResponse(quota);

    const createdAt = new Date().toISOString();
    const jobsBucket = getJobsBucket(env);
    const keys = [
      `jobs/${jobId}/owner.json`,
      `jobs/${jobId}/status.json`,
      ...files.map((file) => `jobs/${jobId}/plays/${file.name}`),
    ];

    try {
      await createJobOwner(env, jobId, user.userId, createdAt);
      for (const file of files) {
        await jobsBucket.put(`jobs/${jobId}/plays/${file.name}`, file, {
          httpMetadata: { contentType: "image/png" },
        });
      }
      await putJson({ PLAYBOOK_BUCKET: jobsBucket }, `jobs/${jobId}/status.json`, {
        status: "processing",
        createdAt,
        mode: "images",
        options: selected,
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
    console.error("Generate error:", err);
    return jsonNoStore({ error: "Internal server error" }, { status: 500 });
  }
}
