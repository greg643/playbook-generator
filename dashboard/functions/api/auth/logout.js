import { jsonNoStore, clearSessionCookie } from "../../_lib/auth.js";

export async function onRequestPost() {
  return jsonNoStore({ ok: true }, { headers: { "Set-Cookie": clearSessionCookie() } });
}
