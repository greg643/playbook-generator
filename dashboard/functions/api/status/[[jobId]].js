export async function onRequestGet(context) {
  const { env, params } = context;
  const jobId = params.jobId;

  if (!jobId) {
    return Response.json({ error: "Missing job ID" }, { status: 400 });
  }

  try {
    const obj = await env.PLAYBOOK_BUCKET.get(`${jobId}/status.json`);

    if (!obj) {
      return Response.json({ error: "Job not found" }, { status: 404 });
    }

    const status = await obj.json();
    return Response.json(status, {
      headers: { "Cache-Control": "no-cache" },
    });
  } catch (err) {
    console.error("Status check error:", err);
    return Response.json(
      { error: "Internal server error" },
      { status: 500 }
    );
  }
}
