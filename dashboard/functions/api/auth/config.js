import { appleConfigured } from "../../_lib/apple.js";
import { jsonNoStore } from "../../_lib/auth.js";
export function onRequestGet({ request, env }) {
  return jsonNoStore({ apple: appleConfigured(env, request) });
}
