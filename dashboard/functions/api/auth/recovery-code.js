import {
  jsonNoStore,
  requireUser,
  createRecoveryFields,
  hasRecentAuthentication,
  emailKey,
} from "../../_lib/auth.js";
import { putJsonIfCurrent } from "../../_lib/r2.js";

export async function onRequestPost(context) {
  const { env } = context;

  try {
    const { user, response } = await requireUser(context);
    if (!user) return response;
    if (user.account.appleAccountKey && typeof user.account.hash !== "string") {
      return jsonNoStore({ error: "This Apple-only account has no password to recover. Use Sign in with Apple." }, { status: 409 });
    }

    if (!hasRecentAuthentication(user)) {
      return jsonNoStore(
        { error: "Please sign in again before replacing your recovery code" },
        { status: 403 }
      );
    }

    const record = { ...user.account };
    const { recoveryCode, fields } = await createRecoveryFields();
    Object.assign(record, fields);
    // Rotating the offline credential is also an explicit revocation of any
    // password-reset link that may still be sitting in an inbox.
    delete record.passwordReset;
    record.recoveryChangedAt = new Date().toISOString();

    const key = await emailKey(user.email);
    const updated = await putJsonIfCurrent(env, key, record, user.accountObject);
    if (updated === null) {
      return jsonNoStore(
        { error: "Account changed while replacing the recovery code; please try again" },
        { status: 409 }
      );
    }

    return jsonNoStore({ recoveryCode });
  } catch (err) {
    console.error("Recovery-code error:", err);
    return jsonNoStore({ error: "Internal server error" }, { status: 500 });
  }
}
