import { jsonNoStore, getUser } from "../../_lib/auth.js";

export async function onRequestGet(context) {
  const { request, env } = context;

  try {
    const user = await getUser(request, env);
    if (!user) {
      return jsonNoStore({ error: "Not signed in" }, { status: 401 });
    }
    return jsonNoStore({ email: user.email, userId: user.userId,
      ...(user.account.appleAccountKey ? { displayName: user.displayName, authMethod: user.authMethod,
        appleLinked: true, passwordAvailable: typeof user.account.hash === "string" } : {}) });
  } catch (err) {
    console.error("Me error:", err);
    return jsonNoStore({ error: "Internal server error" }, { status: 500 });
  }
}
