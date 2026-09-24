# Sign in with Apple

Implemented behind a production-only configuration gate. **This document is not
proof of activation**: the Apple return URL, Pages secrets/bindings and real-device
acceptance tests must all be completed before advertising Apple sign-in.

## Identity and migration

- Use the GSS Services ID `com.greenwichsports.gss.web`, associated with the same
  primary App ID as the native app (`com.greenwichsports.gss`). Do not create an
  unrelated Services ID: a matching Apple email does not establish shared identity.
- Pages exchanges a one-use authorization code directly with Apple. The existing
  GSS `/v2/auth/apple` endpoint verifies the ID token and hashed nonce and supplies
  the same `account_key` used by the native app. It must accept the web Services ID
  as an audience. Browser-provided ID tokens and email addresses are not trusted.
- Existing coaches explicitly authenticate both Apple and their old playbook
  account. Linking retains the immutable local user ID and all existing playbooks;
  it preserves the password/recovery hashes, removes outstanding reset tokens,
  increments the session version and signs out older sessions. No automatic email merge.
- Apple is optional: new email/password registration and password reset remain
  available even with the flag enabled. Linked email accounts can use either login.
  New Apple-only accounts store no email or password and use Apple for recovery;
  adding an email/password to an Apple-only account is not implemented.
- This is shared identity, **not shared playbook storage** or cross-site session
  cookies. App browsing/PDF sync is a separate project. Never share the Pages
  session-signing secret with the GSS app just to make sessions look the same.

## Production setup

1. In Apple Developer → Certificates, Identifiers & Profiles → Services IDs,
   edit the existing GSS web Services ID's Sign in with Apple configuration.
   Add `playbook-generator.pages.dev` to domains and register the exact return URL:
   `https://playbook-generator.pages.dev/api/auth/apple/callback`.
   Keep every existing GSS domain and return URL. Follow Apple's current web
   configuration instructions before enabling the flag.
2. Set these **production** Pages variables/secrets, preserving existing bindings:

   | Name | Value |
   | --- | --- |
   | `PUBLIC_ORIGIN` | `https://playbook-generator.pages.dev` (no trailing slash) |
   | `APPLE_SERVICES_ID` | `com.greenwichsports.gss.web` |
   | `APPLE_TEAM_ID` | Existing GSS Apple Developer Team ID |
   | `APPLE_KEY_ID` | ID of the Sign in with Apple key authorized for GSS |
   | `APPLE_PRIVATE_KEY` | That key's PKCS#8 `.p8` contents, stored as a secret |
   | `GSS_API_BASE` | Existing HTTPS GSS backend origin/base path, not the website |
   | `SESSION_SECRET` | Current Playbook Editor signing secret, exactly 64 hex characters |
   | `APPLE_ENABLED` | Keep `false` until the other setup and verification is complete |

   Use `wrangler pages secret put NAME --project-name playbook-generator` with
   a secure prompt/file input. Never echo, log, commit or pass keys as arguments.
   If the existing signing secret lives at R2 `auth/secret`, copy that exact value
   securely; generating a replacement signs out all existing users.
3. Keep the existing `PLAYBOOK_BUCKET` binding unchanged. Create a small dedicated
   R2 bucket (suggested name: `gss-playbook-auth-state`) with a one-day lifecycle and
   bind it as `AUTH_STATE_BUCKET`. It stores only one-use `apple-once/` replay
   markers. **Never apply its expiry rule to account storage.** This separates
   Apple setup from the existing PDF-job storage and requires no Actions changes.
4. Deploy the reviewed code. Keep Apple disabled on preview hosts and never give
   previews production account storage or private Apple keys. The configuration
   gate also rejects any hostname other than `PUBLIC_ORIGIN`.
5. Enable `APPLE_ENABLED=true`, redeploy, then run the acceptance checks below.
   `/api/auth/config` returns only `{ "apple": true }`, never credentials.

Apple setup guidance: https://developer.apple.com/help/account/capabilities/configure-sign-in-with-apple-for-the-web

`dashboard/apple-continue.png` is Apple's generated English button artwork from
`https://appleid.cdn-apple.com/appleid/button?type=continue&color=white&border=false&border_radius=8&width=330&height=48&locale=en_US`.
It is hosted locally so the sign-in page does not send an image request to Apple
before the coach chooses Apple sign-in.

## Acceptance checks

- Real Apple sign-in on Safari/iPad and desktop: cancellation, retry, first-time
  account, repeat sign-in and converter return path. Confirm the returned GSS
  account key matches the same coach in the app without logging any tokens.
- Use a disposable email account with saved plays. Link it, confirm all books
  remain unchanged, then verify both Apple and email/password can sign in to the
  same account. The old session and pending reset link must stop working, but the
  existing password/recovery code must remain usable.
- Try linking after the local session is older than ten minutes: the existing
  account button must request password authentication again (no redirect loop).
- Export a playbook, choose included plays, preview, generate and download coach
  PDFs; ensure a second account cannot read those plays or downloads.
- Test Apple-only and migrated-account deletion, including a retry after a
  transient storage failure. Deletion requires recent Apple auth + typed DELETE.
- Email accounts can still sign in, create accounts and reset passwords. The
  email service remains necessary; optional Apple sign-in does not replace it.

The September 22 read-only production inspection found only `PLAYBOOK_BUCKET`,
not the separate `JOBS_BUCKET` described in the main README. That pre-existing
job-storage isolation work remains important before wider rollout, but this
Apple feature does not alter or migrate live PDF jobs.

## Boundaries and rollback

Apple authentication is delegated; the editor still owns access control and its
signed, 30-day sessions. Per-account revocation versions are checked on every
authenticated request. Apple server-to-server revocation notifications are not
implemented here, so withdrawing Apple consent does not instantly revoke an
already-issued editor session. A complete GSS identity lifecycle (global logout,
Apple token revocation/deletion and app-wide account deletion) should be handled
centrally; this editor does not revoke the GSS-wide Apple grant when deleting
only its own playbooks.

Keep the new auth-reading code during rollback. After migration, old code cannot
read Apple-only accounts. Turning Apple off prevents Apple-only coaches from
signing back in; linked email accounts retain their password sign-in.
Prefer repairing Apple configuration or retaining the auth endpoints while
rolling back unrelated UI changes. Never restore password hashes from a backup
as an automatic rollback strategy.

Automated tests: `node --test tests/api/apple_auth.test.mjs` covers sealed proof
expiry/origin/tampering, state/nonce and one-use callbacks, code exchange, explicit
linking, collisions, compare-and-swap retries, session revocation and deletion.
Mocks do not replace the real-device acceptance checks above.
