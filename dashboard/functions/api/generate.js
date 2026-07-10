import { jsonNoStore, requireUser } from "../_lib/auth.js";

const OFFENSE_NAME_RE = /^(0[1-9]|1[0-6])\.png$/;
const DEFENSE_NAME_RE = /^D[1-6]\.png$/;
const MAX_FILE_BYTES = 4 * 1024 * 1024;
const MAX_TOTAL_BYTES = 60 * 1024 * 1024;

export async function onRequestPost(context) {
  const { request, env } = context;

  try {
    const { user, response } = await requireUser(context);
    if (!user) return response;

    const contentType = request.headers.get("content-type") || "";
    if (!contentType.includes("multipart/form-data")) {
      return jsonNoStore({ error: "Expected multipart/form-data" }, { status: 400 });
    }

    const formData = await request.formData();

    let options;
    try {
      options = JSON.parse(formData.get("options"));
    } catch (err) {
      return jsonNoStore({ error: "Invalid options" }, { status: 400 });
    }
    // JSON.parse(null) === null without throwing; also reject arrays/scalars.
    if (!options || typeof options !== "object" || Array.isArray(options)) {
      return jsonNoStore({ error: "Invalid options" }, { status: 400 });
    }
    const offense_coach_card = options.offense_coach_card === true;
    const offense_wristband = options.offense_wristband === true;
    const defense_coach_card = options.defense_coach_card === true;
    const defense_wristband = options.defense_wristband === true;

    if (!offense_coach_card && !offense_wristband && !defense_coach_card && !defense_wristband) {
      return jsonNoStore({ error: "Select at least one output" }, { status: 400 });
    }

    const files = formData.getAll("plays");
    if (files.length === 0) {
      return jsonNoStore({ error: "No play images uploaded" }, { status: 400 });
    }

    let totalBytes = 0;
    let hasOffense = false;
    let hasDefense = false;
    for (const file of files) {
      if (!(file instanceof File)) {
        return jsonNoStore({ error: "Invalid play upload" }, { status: 400 });
      }
      if (OFFENSE_NAME_RE.test(file.name)) {
        hasOffense = true;
      } else if (DEFENSE_NAME_RE.test(file.name)) {
        hasDefense = true;
      } else {
        return jsonNoStore({ error: `Invalid play file name: ${file.name}` }, { status: 400 });
      }
      if (file.type !== "image/png") {
        return jsonNoStore({ error: "Play images must be PNG" }, { status: 400 });
      }
      if (file.size > MAX_FILE_BYTES) {
        return jsonNoStore({ error: `Play image too large (max 4 MB): ${file.name}` }, { status: 400 });
      }
      totalBytes += file.size;
    }
    if (totalBytes > MAX_TOTAL_BYTES) {
      return jsonNoStore({ error: "Total upload too large (max 60 MB)" }, { status: 400 });
    }
    if ((offense_coach_card || offense_wristband) && !hasOffense) {
      return jsonNoStore({ error: "Offense outputs selected but no offense plays" }, { status: 400 });
    }
    if ((defense_coach_card || defense_wristband) && !hasDefense) {
      return jsonNoStore({ error: "Defense outputs selected but no defense plays" }, { status: 400 });
    }

    const jobId = crypto.randomUUID();

    for (const file of files) {
      await env.PLAYBOOK_BUCKET.put(`jobs/${jobId}/plays/${file.name}`, await file.arrayBuffer(), {
        httpMetadata: { contentType: "image/png" },
      });
    }

    await env.PLAYBOOK_BUCKET.put(
      `jobs/${jobId}/status.json`,
      JSON.stringify({
        status: "processing",
        createdAt: new Date().toISOString(),
        mode: "images",
        options: { offense_coach_card, offense_wristband, defense_coach_card, defense_wristband },
      }),
      { httpMetadata: { contentType: "application/json" } }
    );

    // Trigger GitHub Actions workflow
    const ghResponse = await fetch(
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

    if (!ghResponse.ok) {
      const errorText = await ghResponse.text();
      console.error("GitHub dispatch failed:", ghResponse.status, errorText);
      await env.PLAYBOOK_BUCKET.put(
        `jobs/${jobId}/status.json`,
        JSON.stringify({ status: "error", message: "Failed to start processing" }),
        { httpMetadata: { contentType: "application/json" } }
      );
      return jsonNoStore({ error: "Failed to trigger processing" }, { status: 502 });
    }

    return jsonNoStore({ jobId });
  } catch (err) {
    console.error("Generate error:", err);
    return jsonNoStore({ error: "Internal server error" }, { status: 500 });
  }
}
