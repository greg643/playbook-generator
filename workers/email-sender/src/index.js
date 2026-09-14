const APP_ORIGIN = "https://playbook-generator.pages.dev";
const FROM_ADDRESS = "no-reply@greenwichsportssystems.com";
const FROM_NAME = "GSS Playbook Editor";
const PASSWORD_RESET_PATH = "/password-reset";
const PERMIT_PATH = "/password-reset/permit";
const MAX_BODY_BYTES = 1024;
const EMAIL_PATTERN = /^[^\s@]+@[^\s@]+\.[^\s@]+$/;
const RESET_TOKEN_PATTERN = /^[A-Za-z0-9_-]{43}$/;
const RATE_LIMIT_KEY_PATTERN = /^[a-f0-9]{64}$/;
const EMAIL_RETRY_DELAYS_MS = [100, 300];
const TRANSIENT_EMAIL_ERROR_CODES = new Set([
  "E_DELIVERY_FAILED",
  "E_INTERNAL_SERVER_ERROR",
  "E_RATE_LIMIT_EXCEEDED",
]);

/** @typedef {(milliseconds: number) => Promise<void>} Sleep */

class HttpError extends Error {
  /**
   * @param {number} status
   * @param {string} message
   */
  constructor(status, message) {
    super(message);
    this.name = "HttpError";
    this.status = status;
  }
}

/**
 * @param {unknown} value
 * @returns {value is Record<string, unknown>}
 */
function isJsonObject(value) {
  return (
    typeof value === "object" &&
    value !== null &&
    !Array.isArray(value) &&
    Object.getPrototypeOf(value) === Object.prototype
  );
}

/**
 * @param {Record<string, unknown>} value
 * @param {string[]} expected
 */
function hasExactKeys(value, expected) {
  const actual = Object.keys(value);
  return (
    actual.length === expected.length &&
    expected.every((key) => Object.hasOwn(value, key))
  );
}

/**
 * Read at most MAX_BODY_BYTES without buffering an unbounded request body.
 *
 * @param {Request} request
 */
async function readBoundedText(request) {
  const contentLength = request.headers.get("content-length");
  if (contentLength !== null) {
    if (!/^\d+$/.test(contentLength)) {
      throw new HttpError(400, "Invalid Content-Length");
    }
    if (Number(contentLength) > MAX_BODY_BYTES) {
      throw new HttpError(413, "Request too large");
    }
  }

  if (request.body === null) return "";

  const reader = request.body.getReader();
  const decoder = new TextDecoder("utf-8", { fatal: true, ignoreBOM: false });
  let totalBytes = 0;
  let text = "";

  try {
    while (true) {
      const { done, value } = await reader.read();
      if (done) break;

      totalBytes += value.byteLength;
      if (totalBytes > MAX_BODY_BYTES) {
        await reader.cancel().catch(() => undefined);
        throw new HttpError(413, "Request too large");
      }
      text += decoder.decode(value, { stream: true });
    }
    return text + decoder.decode();
  } catch (error) {
    if (error instanceof HttpError) throw error;
    throw new HttpError(400, "Body must be valid UTF-8");
  } finally {
    reader.releaseLock();
  }
}

/**
 * @param {Request} request
 * @returns {Promise<Record<string, unknown>>}
 */
async function readJsonObject(request) {
  const contentType = request.headers.get("content-type");
  const mediaType = contentType?.split(";", 1)[0].trim().toLowerCase();
  if (mediaType !== "application/json") {
    throw new HttpError(415, "Content-Type must be application/json");
  }

  const bodyText = await readBoundedText(request);
  /** @type {unknown} */
  let body;
  try {
    body = JSON.parse(bodyText);
  } catch {
    throw new HttpError(400, "Invalid JSON");
  }

  if (!isJsonObject(body)) {
    throw new HttpError(400, "JSON body must be an object");
  }
  return body;
}

/**
 * Keep this in lockstep with the Pages account validator so an address that
 * can own an account can also receive its password-reset email. The explicit
 * CR/LF check documents the header-injection boundary even though \s rejects
 * both characters too.
 *
 * @param {unknown} value
 * @returns {value is string}
 */
function isValidEmail(value) {
  return (
    typeof value === "string" &&
    value.length <= 254 &&
    !/[\r\n]/.test(value) &&
    EMAIL_PATTERN.test(value)
  );
}

/**
 * @param {string} value
 */
function escapeHtml(value) {
  return value.replace(/[&<>"']/g, (character) => {
    switch (character) {
      case "&":
        return "&amp;";
      case "<":
        return "&lt;";
      case ">":
        return "&gt;";
      case '"':
        return "&quot;";
      default:
        return "&#39;";
    }
  });
}

/**
 * @param {string} email
 * @param {string} token
 */
function resetUrl(email, token) {
  return (
    `${APP_ORIGIN}/reset-password#email=${encodeURIComponent(email)}` +
    `&token=${encodeURIComponent(token)}`
  );
}

/**
 * @param {string} url
 */
function emailText(url) {
  return `Reset your GSS Playbook Editor password

A password reset was requested for your GSS Playbook Editor account.

Reset your password:
${url}

For your security, this link expires soon and can be used only once. If the link has expired, request a new one from the sign-in screen.

If you did not request this reset, you can ignore this email. Your password has not been changed.`;
}

/**
 * @param {string} url
 */
function emailHtml(url) {
  const safeUrl = escapeHtml(url);
  return `<!doctype html>
<html lang="en">
  <body style="margin:0;background:#f3f7f2;font-family:Arial,Helvetica,sans-serif;color:#183024">
    <div style="display:none;max-height:0;overflow:hidden">Reset your GSS Playbook Editor password.</div>
    <table role="presentation" width="100%" cellspacing="0" cellpadding="0" style="background:#f3f7f2;padding:28px 12px">
      <tr>
        <td align="center">
          <table role="presentation" width="100%" cellspacing="0" cellpadding="0" style="max-width:560px;background:#ffffff;border:1px solid #d9e6dc;border-radius:12px">
            <tr>
              <td style="padding:32px">
                <p style="margin:0 0 8px;font-size:14px;font-weight:700;color:#237044">GSS PLAYBOOK EDITOR</p>
                <h1 style="margin:0 0 18px;font-size:26px;line-height:1.25;color:#183024">Reset your password</h1>
                <p style="margin:0 0 22px;font-size:16px;line-height:1.55">A password reset was requested for your account.</p>
                <p style="margin:0 0 24px">
                  <a href="${safeUrl}" style="display:inline-block;background:#23834b;color:#ffffff;text-decoration:none;font-size:16px;font-weight:700;padding:13px 20px;border-radius:8px">Reset password</a>
                </p>
                <p style="margin:0 0 18px;font-size:14px;line-height:1.55;color:#44564c">For your security, this link expires soon and can be used only once. If it has expired, request a new one from the sign-in screen.</p>
                <p style="margin:0 0 8px;font-size:13px;line-height:1.5;color:#607066">If the button does not work, copy and paste this address into your browser:</p>
                <p style="margin:0 0 22px;font-size:13px;line-height:1.5;word-break:break-all"><a href="${safeUrl}" style="color:#237044">${safeUrl}</a></p>
                <p style="margin:0;font-size:14px;line-height:1.55;color:#44564c">If you did not request this reset, you can ignore this email. Your password has not been changed.</p>
              </td>
            </tr>
          </table>
        </td>
      </tr>
    </table>
  </body>
</html>`;
}

/**
 * @param {unknown} error
 */
function safeErrorCode(error) {
  if (typeof error !== "object" || error === null) return "UNKNOWN";
  const code = Reflect.get(error, "code");
  return typeof code === "string" && /^[A-Z0-9_]{1,64}$/.test(code)
    ? code
    : "UNKNOWN";
}

/**
 * @param {number} milliseconds
 * @returns {Promise<void>}
 */
function sleep(milliseconds) {
  return new Promise((resolve) => {
    setTimeout(resolve, milliseconds);
  });
}

/**
 * @param {unknown} body
 * @param {number} status
 * @param {HeadersInit} [headers]
 */
function json(body, status = 200, headers = undefined) {
  const responseHeaders = new Headers(headers);
  responseHeaders.set("cache-control", "no-store");
  responseHeaders.set("content-type", "application/json; charset=utf-8");
  responseHeaders.set("x-content-type-options", "nosniff");
  return new Response(JSON.stringify(body), { status, headers: responseHeaders });
}

/**
 * @param {Request} request
 * @param {Env} env
 * @param {Sleep} wait
 */
async function sendPasswordReset(request, env, wait) {
  const body = await readJsonObject(request);
  if (!hasExactKeys(body, ["email", "token"])) {
    throw new HttpError(400, "Expected exactly email and token");
  }

  const { email, token } = body;
  if (!isValidEmail(email) || typeof token !== "string" || !RESET_TOKEN_PATTERN.test(token)) {
    throw new HttpError(400, "Invalid email or token");
  }

  const url = resetUrl(email, token);
  const message = {
    to: email,
    from: { email: FROM_ADDRESS, name: FROM_NAME },
    subject: "Reset your GSS Playbook Editor password",
    text: emailText(url),
    html: emailHtml(url),
  };

  for (let attempt = 0; attempt <= EMAIL_RETRY_DELAYS_MS.length; attempt += 1) {
    try {
      await env.EMAIL.send(message);
      return json({ ok: true });
    } catch (error) {
      const code = safeErrorCode(error);
      const canRetry =
        TRANSIENT_EMAIL_ERROR_CODES.has(code) && attempt < EMAIL_RETRY_DELAYS_MS.length;
      if (canRetry) {
        await wait(EMAIL_RETRY_DELAYS_MS[attempt]);
        continue;
      }

      console.error(JSON.stringify({
        event: "password_reset_email_failed",
        code,
        attempts: attempt + 1,
      }));
      return json({ error: "Unable to send password reset email" }, 502);
    }
  }

  // The loop always returns after a successful send or the final failure.
  return json({ error: "Unable to send password reset email" }, 502);
}

/**
 * @param {Request} request
 * @param {Env} env
 */
async function permitPasswordReset(request, env) {
  const body = await readJsonObject(request);
  if (!hasExactKeys(body, ["key"])) {
    throw new HttpError(400, "Expected exactly key");
  }

  const { key } = body;
  if (typeof key !== "string" || !RATE_LIMIT_KEY_PATTERN.test(key)) {
    throw new HttpError(400, "Invalid rate-limit key");
  }

  try {
    const { success } = await env.RESET_RATE_LIMITER.limit({ key });
    return json({ allowed: success });
  } catch (error) {
    console.error(JSON.stringify({
      event: "password_reset_rate_limit_failed",
      code: safeErrorCode(error),
    }));
    return json({ allowed: false }, 503);
  }
}

/**
 * @param {Request} request
 * @param {Env} env
 * @param {Sleep} [wait]
 */
export async function handleRequest(request, env, wait = sleep) {
  const { pathname } = new URL(request.url);
  if (pathname !== PASSWORD_RESET_PATH && pathname !== PERMIT_PATH) {
    return json({ error: "Not found" }, 404);
  }

  if (request.method !== "POST") {
    return json({ error: "Method not allowed" }, 405, { allow: "POST" });
  }

  try {
    return pathname === PASSWORD_RESET_PATH
      ? await sendPasswordReset(request, env, wait)
      : await permitPasswordReset(request, env);
  } catch (error) {
    if (error instanceof HttpError) {
      return json({ error: error.message }, error.status);
    }

    console.error(JSON.stringify({
      event: "password_reset_worker_failed",
      code: safeErrorCode(error),
    }));
    return json({ error: "Internal server error" }, 500);
  }
}

/** @type {ExportedHandler<Env>} */
const worker = {
  async fetch(request, env) {
    return handleRequest(request, env);
  },
};

export default worker;
