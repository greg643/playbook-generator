import {
  emailKey,
  isValidEmail,
  jsonNoStore,
  normalizeEmail,
  readBoundedUtf8Text,
  sha256Hex,
} from "../../../_lib/auth.js";
import {
  createPasswordResetChallenge,
  nextPasswordResetRate,
  PASSWORD_RESET_GENERIC_MESSAGE,
  sendPasswordResetEmailSafely,
} from "../../../_lib/password-reset.js";
import { putJsonIfCurrent } from "../../../_lib/r2.js";

const MAX_BODY_BYTES = 2048;

function genericResponse() {
  return jsonNoStore({ message: PASSWORD_RESET_GENERIC_MESSAGE }, { status: 202 });
}

function unavailableResponse() {
  return jsonNoStore(
    {
      error:
        "Email password reset is temporarily unavailable. Use your recovery code or try again later.",
    },
    { status: 503 }
  );
}

async function requestAllowed(request, env) {
  if (!env.EMAIL_SERVICE || typeof env.EMAIL_SERVICE.fetch !== "function") {
    throw new Error("Missing EMAIL_SERVICE binding");
  }
  const clientAddress = request.headers.get("cf-connecting-ip") || "unknown";
  const key = await sha256Hex(`gss-password-reset-rate:v1:${clientAddress}`);
  const response = await env.EMAIL_SERVICE.fetch("https://email.internal/password-reset/permit", {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify({ key }),
  });
  if (!response.ok) throw new Error(`Email service permit check returned ${response.status}`);
  const result = await response.json().catch(() => null);
  if (!result || typeof result.allowed !== "boolean") {
    throw new Error("Email service returned an invalid permit response");
  }
  return result.allowed;
}

function activeAccountRecord(record, email) {
  return Boolean(
    record &&
      typeof record.userId === "string" &&
      normalizeEmail(record.email) === email &&
      !record.deletedAt &&
      !record.disabledAt
  );
}

async function challengeWasCommitted(env, key, requestId, tokenHash) {
  const object = await env.PLAYBOOK_BUCKET.get(key);
  if (!object) return false;
  try {
    const record = await object.json();
    return Boolean(
      record &&
        record.passwordReset &&
        record.passwordReset.requestId === requestId &&
        record.passwordReset.tokenHash === tokenHash
    );
  } catch (error) {
    return false;
  }
}

async function issueAndSendPasswordReset(env, email) {
  try {
    const key = await emailKey(email);
    for (let attempt = 0; attempt < 3; attempt += 1) {
      const object = await env.PLAYBOOK_BUCKET.get(key);
      if (!object) return;

      let record;
      try {
        record = await object.json();
      } catch (error) {
        return;
      }
      if (!activeAccountRecord(record, email)) return;

      const nowMs = Date.now();
      const rate = nextPasswordResetRate(record, nowMs);
      if (!rate) return;
      const { token, challenge } = await createPasswordResetChallenge(record, nowMs);
      const updatedRecord = {
        ...record,
        passwordReset: challenge,
        passwordResetRate: rate,
      };

      let updated;
      try {
        updated = await putJsonIfCurrent(env, key, updatedRecord, object);
      } catch (error) {
        if (
          await challengeWasCommitted(
            env,
            key,
            challenge.requestId,
            challenge.tokenHash
          )
        ) {
          await sendPasswordResetEmailSafely(env, email, token);
          return;
        }
        // A transient failure before the conditional write committed is safe
        // to retry. Each attempt re-reads the account and only emails the
        // challenge proven to be stored.
        continue;
      }
      if (updated !== null) {
        await sendPasswordResetEmailSafely(env, email, token);
        return;
      }
    }
  } catch (error) {
    console.error(JSON.stringify({ event: "password_reset_issue_failed" }));
  }
}

export async function onRequestPost(context) {
  const { request, env } = context;

  try {
    const bodyText = await readBoundedUtf8Text(request, MAX_BODY_BYTES);
    if (bodyText === null) {
      return jsonNoStore({ error: "Request too large" }, { status: 413 });
    }

    let body;
    try {
      body = JSON.parse(bodyText);
    } catch (error) {
      return jsonNoStore({ error: "Invalid JSON" }, { status: 400 });
    }

    let allowed;
    try {
      allowed = await requestAllowed(request, env);
    } catch (error) {
      console.error(
        JSON.stringify({ event: "password_reset_email_service_unavailable" })
      );
      return unavailableResponse();
    }
    if (!allowed) return genericResponse();

    const email = normalizeEmail(body && body.email);
    if (!isValidEmail(email)) return genericResponse();

    // Account lookup, challenge issuance, and delivery all happen after the
    // uniform response path. Missing and existing accounts therefore schedule
    // the same background operation and expose no provider timing difference.
    const delivery = issueAndSendPasswordReset(env, email);
    if (typeof context.waitUntil === "function") {
      context.waitUntil(delivery);
    } else {
      await delivery;
    }
    return genericResponse();
  } catch (error) {
    console.error(JSON.stringify({ event: "password_reset_request_failed" }));
    return jsonNoStore({ error: "Internal server error" }, { status: 500 });
  }
}
