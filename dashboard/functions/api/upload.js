export async function onRequestPost(context) {
  const { request, env } = context;

  try {
    const contentType = request.headers.get("content-type") || "";
    if (!contentType.includes("multipart/form-data")) {
      return Response.json(
        { error: "Expected multipart/form-data" },
        { status: 400 }
      );
    }

    const formData = await request.formData();
    const file = formData.get("file");

    if (!file || !(file instanceof File)) {
      return Response.json({ error: "No file uploaded" }, { status: 400 });
    }

    if (!file.name.endsWith(".pptx")) {
      return Response.json(
        { error: "Only .pptx files are accepted" },
        { status: 400 }
      );
    }

    // 50 MB limit
    if (file.size > 50 * 1024 * 1024) {
      return Response.json(
        { error: "File too large (max 50 MB)" },
        { status: 400 }
      );
    }

    // Generate job ID
    const jobId = crypto.randomUUID();

    // Upload PPTX to R2
    const fileBuffer = await file.arrayBuffer();
    await env.PLAYBOOK_BUCKET.put(`${jobId}/input.pptx`, fileBuffer, {
      httpMetadata: { contentType: "application/vnd.openxmlformats-officedocument.presentationml.presentation" },
    });

    // Write initial status
    await env.PLAYBOOK_BUCKET.put(
      `${jobId}/status.json`,
      JSON.stringify({ status: "processing", createdAt: new Date().toISOString() }),
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
      // Update status to error
      await env.PLAYBOOK_BUCKET.put(
        `${jobId}/status.json`,
        JSON.stringify({ status: "error", message: "Failed to start processing" }),
        { httpMetadata: { contentType: "application/json" } }
      );
      return Response.json(
        { error: "Failed to trigger processing" },
        { status: 502 }
      );
    }

    return Response.json({ jobId });
  } catch (err) {
    console.error("Upload error:", err);
    return Response.json(
      { error: "Internal server error" },
      { status: 500 }
    );
  }
}
