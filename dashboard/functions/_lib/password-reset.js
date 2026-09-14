import {
  accountSessionVersion,
  generatePasswordResetToken,
  hashPasswordResetToken,
} from "./auth.js";

export const PASSWORD_RESET_TTL_MS = 15 * 60 * 1000;
export const PASSWORD_RESET_MIN_INTERVAL_MS = 60 * 1000;
export const PASSWORD_RESET_WINDOW_MS = 60 * 60 * 1000;
export const PASSWORD_RESET_MAX_PER_WINDOW = 3;
export const PASSWORD_RESET_DAY_WINDOW_MS = 24 * 60 * 60 * 1000;
export const PASSWORD_RESET_MAX_PER_DAY = 5;
export const PASSWORD_RESET_GENERIC_MESSAGE =
  "If an active account exists for that email, a password-reset link has been sent.";

function parseTimestamp(value) {
  if (typeof value !== "string") return null;
  const parsed = Date.parse(value);
  return Number.isFinite(parsed) ? parsed : null;
}

// Return the rate record that should be committed with a new challenge, or
// null when this account is still cooling down. This state deliberately lives
// in the credential object so issuing the challenge and consuming quota share
// one R2 compare-and-swap.
export function nextPasswordResetRate(record, nowMs = Date.now()) {
  const current = record && record.passwordResetRate;
  const lastSentAt = parseTimestamp(current && current.lastSentAt);
  if (lastSentAt !== null && nowMs - lastSentAt < PASSWORD_RESET_MIN_INTERVAL_MS) {
    return null;
  }

  const storedWindowStart = parseTimestamp(current && current.windowStartedAt);
  const inCurrentWindow =
    storedWindowStart !== null &&
    storedWindowStart <= nowMs &&
    nowMs - storedWindowStart < PASSWORD_RESET_WINDOW_MS;
  const windowStartedAt = inCurrentWindow ? storedWindowStart : nowMs;
  const previousCount =
    inCurrentWindow && Number.isSafeInteger(current && current.sendCount) && current.sendCount >= 0
      ? current.sendCount
      : 0;
  if (previousCount >= PASSWORD_RESET_MAX_PER_WINDOW) return null;

  const storedDayStart = parseTimestamp(current && current.dayStartedAt);
  const inCurrentDay =
    storedDayStart !== null &&
    storedDayStart <= nowMs &&
    nowMs - storedDayStart < PASSWORD_RESET_DAY_WINDOW_MS;
  const dayStartedAt = inCurrentDay ? storedDayStart : nowMs;
  const previousDayCount =
    inCurrentDay && Number.isSafeInteger(current && current.daySendCount) && current.daySendCount >= 0
      ? current.daySendCount
      : 0;
  if (previousDayCount >= PASSWORD_RESET_MAX_PER_DAY) return null;

  return {
    schema: 1,
    windowStartedAt: new Date(windowStartedAt).toISOString(),
    sendCount: previousCount + 1,
    dayStartedAt: new Date(dayStartedAt).toISOString(),
    daySendCount: previousDayCount + 1,
    lastSentAt: new Date(nowMs).toISOString(),
  };
}

export async function createPasswordResetChallenge(record, nowMs = Date.now()) {
  const token = generatePasswordResetToken();
  return {
    token,
    challenge: {
      schema: 1,
      requestId: crypto.randomUUID(),
      userId: record.userId,
      tokenHash: await hashPasswordResetToken(token),
      issuedForSessionVersion: accountSessionVersion(record),
      issuedAt: new Date(nowMs).toISOString(),
      expiresAt: new Date(nowMs + PASSWORD_RESET_TTL_MS).toISOString(),
    },
  };
}

export function isUsablePasswordResetChallenge(record, nowMs = Date.now()) {
  const challenge = record && record.passwordReset;
  const expiresAt = parseTimestamp(challenge && challenge.expiresAt);
  return Boolean(
    challenge &&
      challenge.schema === 1 &&
      typeof challenge.requestId === "string" &&
      typeof record.userId === "string" &&
      challenge.userId === record.userId &&
      typeof challenge.tokenHash === "string" &&
      /^[0-9a-f]{64}$/i.test(challenge.tokenHash) &&
      challenge.issuedForSessionVersion === accountSessionVersion(record) &&
      expiresAt !== null &&
      expiresAt > nowMs
  );
}

export async function sendPasswordResetEmail(env, email, token) {
  const response = await env.EMAIL_SERVICE.fetch("https://email.internal/password-reset", {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify({ email, token }),
  });
  if (!response.ok) {
    throw new Error(`Email service returned ${response.status}`);
  }
}

export async function sendPasswordResetEmailSafely(env, email, token) {
  try {
    await sendPasswordResetEmail(env, email, token);
  } catch (error) {
    const category =
      error instanceof Error && /^Email service returned [45]\d\d$/.test(error.message)
        ? error.message
        : "Email service unavailable";
    console.error(
      JSON.stringify({
        event: "password_reset_email_failed",
        error: category,
      })
    );
  }
}
