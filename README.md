# Playbook Generator

Web app that converts a flag football PowerPoint playbook (`.pptx`) into printable coach cards and wristband PDFs.

## What it produces

- **Offense Coach Card** — 4x4 grid of all 16 plays
- **Defense Coach Card** — 2x2 grid of defensive formations
- **Offense Wristband** — Cut-and-laminate cards sized for QB wristband holders
- **Defense Wristband** — Cut-and-laminate defense reference cards

## Architecture

```
Browser → Cloudflare Pages (static HTML)
        → Pages Functions (authentication, saves, job ownership and quotas)
        → Account R2 bucket (credentials + saved playbooks)
        → Job R2 bucket (uploads, rendered plays, status + PDFs)
        → GitHub Actions (job-bucket credentials only)
        → Pages Functions /api/status + /api/download (owner checked)
```

Keep the account and job buckets separate in production. User-controlled PPTX
files are parsed by LibreOffice, Poppler, Pillow and python-pptx in the Actions
runner; that runner must not have credentials for password records, saved
playbooks, or the session-signing key.

## Play Editor

Browser-based editor at `/editor` — sign in, build plays directly (drag the fixed player chips, draw routes/lines/labels), and generate the same four PDFs without PowerPoint.

- **Auth**: email + password. Passwords are hashed with PBKDF2-SHA256 (per-user
  salt, 600k iterations; legacy hashes upgrade on login). Sessions are
  HMAC-signed, account/version checked, and revoked after password recovery or
  deletion. Recovery codes are single-use and rotated atomically. Account
  deletion first writes a blocking tombstone and durable job inventory; cleanup
  is idempotent and can be resumed after a transient storage failure without
  re-enabling a partially deleted account. Finalization leaves only a minimal
  conditional credential tombstone, which registration can safely replace if
  the same email is used again.
- **R2 keys**:
  - `auth/secret` — legacy session signing secret fallback; use `SESSION_SECRET`
    in production
  - `users/byemail/<sha256(email)>.json` — credential record (userId, salt, hash)
  - `accounts/<userId>/playbook.json` — the user's saved plays (JSON)
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

### GitHub Actions secrets

- `R2_ENDPOINT`, `R2_ACCESS_KEY_ID`, `R2_SECRET_ACCESS_KEY`, `R2_BUCKET` — access
  to the **job bucket only**
- `CLOUDFLARE_API_TOKEN`, `CLOUDFLARE_ACCOUNT_ID` — Pages deployment

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
5. Configure the variables, secrets and edge rate limits above, then deploy
   `dashboard/`.
6. Verify registration/recovery, cross-account job denial, PPTX and editor-mode
   generation, PDF downloads, quota responses and account deletion in staging.
