# GSS Playbook Editor

Greenwich Sports Systems web app for drawing flag football plays or converting a PowerPoint playbook (`.pptx`) into printable coach cards and wristband PDFs.

## Web routes

- `/` — account sign-in and the two-choice GSS Playbook Editor home
- `/editor` — GSS Playbook Editor for 5v5 and 6v6 plays
- `/pptx-guide` — public PowerPoint (PPTX) Import Guide with non-proprietary deck examples
- `/converter` — signed-in PowerPoint (PPTX) Import screen
- `/reset-password` — single-use emailed password-reset link target
- `/help` — editor, account, compatibility, and printing help

The PowerPoint (PPTX) Import page verifies the current session before revealing its controls. The upload,
status, and download APIs remain the actual security boundary: they require an authenticated
account, and job results are owner-checked.

## What it produces

- **Offense Coach Card** — 4x4 grid, automatically paginated every 16 plays
- **Defense Coach Card** — Full-page grid of up to 6 defensive formations per page
- **Offense Wristband** — Cut-and-laminate cards, automatically paginated every 8 plays
- **Defense Wristband** — Cut-and-laminate defense reference cards

## Architecture

```
Browser → Cloudflare Pages (static HTML)
        → Pages Functions (authentication, saves, job ownership and quotas)
        → Private email Worker (rate gate + fixed password-reset template)
        → Cloudflare Email Service (transactional delivery)
        → Account R2 bucket (credentials + saved playbooks)
        → Job R2 bucket (uploads, rendered plays, status + PDFs)
        → GitHub Actions (job-bucket credentials only)
        → Pages Functions /api/status + /api/download (owner checked)
```

Keep the account and job buckets separate in production. User-controlled PPTX
files are parsed by LibreOffice, Poppler, Pillow and python-pptx in the Actions
runner; that runner must not have credentials for password records, saved
playbooks, or the session-signing key.

## GSS Playbook Editor

The browser-based **GSS Playbook Editor** lives at `/editor`. Sign in at `/`,
choose **Playbook Editor**, drag the player chips into position, draw routes,
lines and labels, and generate the same four PDFs without PowerPoint.

The editor supports both common formats:

- **5v5 offense:** `C`, `1`, `2`, `3` and `QB`; **5v5 defense:** `1`–`5`
- **6v6 offense:** `1`–`5` and `QB`; **6v6 defense:** `1`–`5` and `N`

New plays default to 5v5, and the editor saves the coach's latest selection for
future plays. The format is also saved with each play, so changing the selection
does not alter existing or duplicated diagrams.

Each account can keep up to 20 named playbooks. The selector at the top of the
editor sidebar switches between them, and **New** creates an empty playbook with
a chosen 5v5 or 6v6 starting format. Seasons are intentionally part of the name
in this first version (for example, `Fall 2026 5v5`). **Rename** changes only the
catalog label; every playbook keeps its own plays, numbering, bench, format
preference, backups and generated PDFs. Existing single-playbook accounts are
adopted automatically as `My Playbook` without moving or rewriting their data.
Playbooks created after that original can be deleted by typing their exact name
to confirm; the original can instead be renamed and reused. Export first if the
plays may be needed again. Each JSON backup includes its source playbook name
and can be shared with another coach or imported into another playbook or account.

- **Auth**: email + password. Passwords are hashed with PBKDF2-SHA256 (per-user
  salt, 100k iterations, the Workers Web Crypto maximum; lower-work-factor
  legacy hashes upgrade on login). Sessions are
  HMAC-signed, account/version checked, and revoked after password recovery or
  deletion. Email-reset links expire after 15 minutes, are single-use, store
  only a SHA-256 token digest, and are issued/consumed with R2 compare-and-swap.
  Delivery is limited to one message per minute, three per one-hour accounting
  window, and five per 24-hour accounting window for an account. Recovery codes
  remain an offline fallback and rotate atomically. Account
  deletion first writes a blocking tombstone and durable job inventory; cleanup
  is idempotent and can be resumed after a transient storage failure without
  re-enabling a partially deleted account. Finalization leaves only a minimal
  conditional credential tombstone, which registration can safely replace if
  the same email is used again.
- **R2 keys**:
  - `auth/secret` — legacy session signing secret fallback; use `SESSION_SECRET`
    in production
  - `users/byemail/<sha256(email)>.json` — credential record (userId, salt, hash)
  - `accounts/<userId>/playbook.json` — backward-compatible default playbook
  - `accounts/<userId>/playbooks/catalog.json` — names and immutable playbook IDs
  - `accounts/<userId>/playbooks/items/<playbookId>.json` — additional playbooks
- **Images-mode jobs**: the editor exports each play to PNG (`01.png`–`16.png`
  offense, `D1.png`–`D6.png` defense). The job bucket stores immutable ownership
  metadata, images and status under `jobs/<jobId>/`. The same generator produces
  the PDFs without running LibreOffice.

## Local development

```bash
python3.11 -m venv .venv
source .venv/bin/activate
pip install -r requirements.txt
python pipeline/playbook_pipeline.py <playbook.pptx> [output_dir]
```

Requires Python 3.11+, LibreOffice and Poppler (`pdftoppm`). Every run uses an
isolated directory under `_playbook_work/`; artifacts from an older deck are
never reused. The pipeline rejects unsafe/oversized PPTX archives, excessive
slide/play counts and images with unsafe dimensions.

### PowerPoint (PPTX) Import behavior

Public PPTX-to-PDF conversion is deterministic; it does not use an LLM to
interpret play marks or player positions. The pipeline identifies the largest
top-level PowerPoint rectangle-family shape on each play slide (including
standard, rounded, and snipped-corner rectangles), renders the visible slide,
and crops that rendering to the field. Because it preserves the drawing rather
than counting or relabeling players, it accepts both 5v5 and 6v6 decks without a
separate format setting. If a deck has no recognized `OFFENSE` divider, valid
play slides before its first `DEFENSE` divider are treated as offense; a
successful conversion warns when it uses this fallback or skips likely plays.
A fully headerless deck remains offense-only. For explicitly sectioned decks,
use nearly empty `OFFENSE` and `DEFENSE` divider slides and start with the
appropriate divider. PowerPoint (PPTX) Import
supports up to 64 offense plays (16 per coach-card page and 8 per wristband page)
and 24 defense plays (6 per coach-card page and 8 per wristband page). See the
[PowerPoint (PPTX) Import Guide](dashboard/pptx-guide.html) for example slides and known
failure cases.

Convert a supported PPTX into an editor JSON backup with:

```bash
python pipeline/pptx_to_editor.py <playbook.pptx> <playbook.json>
```

## Tests

```bash
python -m unittest discover -s tests/pipeline -v
node --test tests/api/*.test.mjs tests/frontend/*.test.mjs tests/workflows/*.test.mjs
```

CI runs these tests plus Python compilation and JavaScript syntax checks on
every push and pull request.

## Production configuration

### Cloudflare Pages bindings and variables

- `PLAYBOOK_BUCKET` — permanent account-data R2 bucket
- `JOBS_BUCKET` — separate transient job R2 bucket
- `EMAIL_SERVICE` — production-only Service binding to the private
  `gss-playbook-email` Worker
- `SESSION_SECRET` — exactly 64 hexadecimal characters; create with
  `openssl rand -hex 32`
- `GITHUB_TOKEN` — token allowed to dispatch this repository's workflow
- `GITHUB_REPO` — optional `owner/repository` override
- `MAX_ACTIVE_JOBS_PER_USER` — optional, default `2`
- `MAX_DAILY_JOBS_PER_USER` — optional, default `20`
- `JOB_ACTIVE_MINUTES` — optional, default `30`
- `JOB_STALE_MINUTES` — optional terminal timeout for stuck status records,
  default `15` (keep above the 10-minute Actions timeout)

Before public traffic, configure Cloudflare edge rate-limit rules for
registration, login, password recovery and account deletion by IP. Also enforce
request-body limits at the edge for `/api/upload` (52 MB) and `/api/generate`
(62 MB): Pages Functions must parse multipart bodies before they can inspect decoded fields.
Application validation and per-account quotas remain the second layer; the
edge rules are a release requirement, not an optional tuning step.

### Password-reset email

Outbound reset messages use Cloudflare Email Service through the route-less
Worker in `workers/email-sender/`. The Worker has no `workers.dev` or preview
URL, accepts only its two service-binding endpoints, owns a coarse rate-limit
binding, constructs links from the fixed production origin, and restricts the
email binding to `no-reply@greenwichsportssystems.com`. Pages never holds an
email API key and cannot choose the sender, subject, HTML, or link origin.
Pages schedules delivery after its uniform reset response, and the Worker
retries documented transient delivery failures three times. This lightweight
path fits the pilot; move delivery to a Queue/outbox before materially higher
volume or stronger delivery guarantees are required.

Email Sending requires the account-wide Workers Paid plan. Onboard and verify
`greenwichsportssystems.com` in Cloudflare Email Sending before deployment so
Cloudflare can install aligned SPF, DKIM, bounce-routing, and DMARC records.
This service is for transactional messages only, not marketing email.

### GitHub Actions secrets

- `R2_ENDPOINT`, `R2_ACCESS_KEY_ID`, `R2_SECRET_ACCESS_KEY`, `R2_BUCKET` — access
  to the **job bucket only**
- `CLOUDFLARE_API_TOKEN`, `CLOUDFLARE_ACCOUNT_ID` — Pages and private Worker
  deployment; the token needs Pages deployment and Workers Scripts edit access

### Migration/deployment order

1. If upgrading an existing deployment, copy the current `auth/secret` value
   into `SESSION_SECRET` exactly. Using a different value signs every user out.
2. Create a separate job bucket and bind it as `JOBS_BUCKET`.
3. Replace the Actions R2 token with credentials scoped only to that job bucket.
4. Add a one-day lifecycle rule to the whole job bucket. Account deletion keeps
   small cancellation/ownership tombstones long enough to stop an already
   dispatched worker, while immediately scrubbing uploaded files and PDFs. Its
   bounded cleanup inventory checks today and yesterday, so this lifecycle rule
   is required for older job payloads. If temporarily using a shared bucket,
   scope expiry to `jobs/` only—never `auth/`, `users/` or `accounts/`.
5. Enable Workers Paid, onboard `greenwichsportssystems.com` in Email Sending,
   and wait for all sending-domain DNS checks to pass.
6. Deploy `workers/email-sender/`, then add the production Pages Service binding
   `EMAIL_SERVICE` targeting `gss-playbook-email`. Keep production email
   unbound from preview deployments.
7. Configure the remaining variables, secrets and edge rate limits above, then
   deploy `dashboard/`.
8. Verify registration/email recovery/offline-code recovery, cross-account job
   denial, PPTX and editor-mode generation, PDF downloads, quota responses and
   account deletion in staging.
