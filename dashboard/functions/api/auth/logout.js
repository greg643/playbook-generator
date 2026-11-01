import { jsonNoStore, clearSessionCookie } from "../../_lib/auth.js";
import { clearAppWebSession } from "../../_lib/app-web-session.js";

export async function onRequestPost({ request, env }) {
  if (request.headers.get('origin') !== new URL(request.url).origin) {
    return jsonNoStore({ error: 'Invalid request origin' }, { status: 403 });
  }
  try { await clearAppWebSession(request, env); }
  catch {
    return jsonNoStore({ error: 'Could not finish signing out. Please try again.' }, { status: 503 });
  }
  return jsonNoStore({ ok: true }, { headers: { "Set-Cookie": clearSessionCookie() } });
}
