// Apple-only is the fail-closed default. This server-only opt-in retains the
// tested legacy migration/rollback path; never enable it in production casually.
export function legacyEmailAuthEnabled(env) {
  return env.LEGACY_EMAIL_AUTH_ENABLED === "true";
}

export function appleOnlyResponse() {
  return Response.json({
    error: "Email and password sign-in is no longer available. Continue with Apple.",
    code: "apple_only",
  }, { status: 410, headers: { "Cache-Control": "private, no-store" } });
}
