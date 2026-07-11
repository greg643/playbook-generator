import { jsonNoStore, requireUser } from "../../_lib/auth.js";
import {
  finishJobSlot,
  getJobsBucket,
  isJobOwner,
  JOB_ID_RE,
  PDF_NAME_RE,
} from "../../_lib/jobs.js";

export async function onRequestGet(context) {
  const { env, params } = context;
  const segments = Array.isArray(params.catchall)
    ? params.catchall
    : typeof params.catchall === "string"
      ? params.catchall.split("/")
      : [];

  if (segments.length !== 2) {
    return jsonNoStore(
      { error: "Expected /api/download/{jobId}/{filename}" },
      { status: 400 }
    );
  }

  const [jobId, filename] = segments;
  if (!JOB_ID_RE.test(jobId)) {
    return jsonNoStore({ error: "Invalid job ID" }, { status: 400 });
  }
  if (!PDF_NAME_RE.test(filename)) {
    return jsonNoStore({ error: "Invalid PDF filename" }, { status: 400 });
  }

  try {
    const { user, response } = await requireUser(context);
    if (!user) return response;

    if (!(await isJobOwner(env, jobId, user.userId))) {
      return jsonNoStore({ error: "File not found" }, { status: 404 });
    }

    // A filename is authorized only after the processor lists it in the final
    // status. This prevents access to partial or unexpected objects under a
    // valid job prefix.
    const jobsBucket = getJobsBucket(env);
    const statusObject = await jobsBucket.get(`jobs/${jobId}/status.json`);
    if (!statusObject) {
      return jsonNoStore({ error: "File not found" }, { status: 404 });
    }
    const status = await statusObject.json();
    if (
      !status ||
      status.status !== "complete" ||
      !Array.isArray(status.files) ||
      !status.files.includes(filename)
    ) {
      return jsonNoStore({ error: "File not found" }, { status: 404 });
    }

    const object = await jobsBucket.get(`jobs/${jobId}/${filename}`);
    if (!object) {
      return jsonNoStore({ error: "File not found" }, { status: 404 });
    }

    const finish = finishJobSlot(env, user.userId, jobId).catch((error) => {
      console.error("Could not release completed job slot:", error);
    });
    if (typeof context.waitUntil === "function") context.waitUntil(finish);
    else await finish;

    return new Response(object.body, {
      headers: {
        "Content-Type": "application/pdf",
        "Content-Disposition": `attachment; filename="${filename}"`,
        "Cache-Control": "private, no-store",
        "X-Content-Type-Options": "nosniff",
      },
    });
  } catch (err) {
    console.error("Download error:", err);
    return jsonNoStore({ error: "Internal server error" }, { status: 500 });
  }
}
