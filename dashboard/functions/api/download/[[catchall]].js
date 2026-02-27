export async function onRequestGet(context) {
  const { env, params } = context;

  // params.catchall is an array of path segments: [jobId, filename]
  const segments = params.catchall;
  if (!segments || segments.length < 2) {
    return Response.json(
      { error: "Expected /api/download/{jobId}/{filename}" },
      { status: 400 }
    );
  }

  const jobId = segments[0];
  const filename = segments.slice(1).join("/");

  // Only allow PDF downloads
  if (!filename.endsWith(".pdf")) {
    return Response.json(
      { error: "Only PDF files can be downloaded" },
      { status: 400 }
    );
  }

  try {
    const obj = await env.PLAYBOOK_BUCKET.get(`${jobId}/${filename}`);

    if (!obj) {
      return Response.json({ error: "File not found" }, { status: 404 });
    }

    return new Response(obj.body, {
      headers: {
        "Content-Type": "application/pdf",
        "Content-Disposition": `attachment; filename="${filename}"`,
        "Cache-Control": "public, max-age=3600",
      },
    });
  } catch (err) {
    console.error("Download error:", err);
    return Response.json(
      { error: "Internal server error" },
      { status: 500 }
    );
  }
}
