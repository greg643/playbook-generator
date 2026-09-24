import { jsonNoStore, requireUser } from "../../../_lib/auth.js";
export async function onRequestGet(context) {
  const { user, response } = await requireUser(context, { allowDisabled: true });
  if (!user) return response;
  if (user.authMethod !== "apple") return jsonNoStore({ error: "Please continue with Apple again." }, { status: 403 });
  return jsonNoStore({ userId: user.userId });
}
