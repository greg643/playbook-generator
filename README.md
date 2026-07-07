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
        → CF Worker /api/upload → R2 bucket + GitHub Actions trigger
        → GitHub Actions: runs playbook_pipeline.py, uploads PDFs to R2
        → CF Worker /api/status/[id] → polls R2 for completion
        → CF Worker /api/download/[id]/[file] → serves PDFs from R2
```

## Play Editor

Browser-based editor at `/editor` — sign in, build plays directly (drag the fixed player chips, draw routes/lines/labels), and generate the same four PDFs without PowerPoint.

- **Auth**: email + password. Passwords hashed with PBKDF2-SHA256 (per-user salt, 100k iterations); sessions are HMAC-signed cookies (`pb_session`, 30 days). The signing secret is auto-created in R2 at `auth/secret` on first use — no setup required.
- **R2 keys**:
  - `auth/secret` — session signing secret
  - `users/byemail/<sha256(email)>.json` — credential record (userId, salt, hash)
  - `accounts/<userId>/playbook.json` — the user's saved plays (JSON)
- **Images-mode jobs**: the editor exports each play to PNG (`01.png`–`16.png` offense, `D1.png`–`D6.png` defense) and `POST /api/generate` stores them at `<jobId>/plays/` with `mode: "images"` in status.json. The same GitHub Actions job downloads the PNGs and runs the same `PlaybookGenerator`, producing identical PDFs — no LibreOffice step. Status polling and downloads reuse the existing endpoints.

## Local development

```bash
cd pipeline
python playbook_pipeline.py <playbook.pptx> [output_dir]
```

Requires: Python 3.11+, LibreOffice, poppler-utils (pdftoppm)

## Infrastructure setup

1. **Cloudflare R2 bucket**: `playbook-files`
2. **CF Pages project**: linked to this repo, deploys `dashboard/`
3. **CF Pages bindings**: R2 binding `PLAYBOOK_BUCKET` → `playbook-files`
4. **CF Pages env vars**: `GITHUB_TOKEN` (PAT with `repo` scope)
5. **GitHub secrets**: `R2_ENDPOINT`, `R2_ACCESS_KEY_ID`, `R2_SECRET_ACCESS_KEY`, `R2_BUCKET`, `CLOUDFLARE_API_TOKEN`, `CLOUDFLARE_ACCOUNT_ID`
6. **R2 lifecycle rule**: auto-delete objects older than 1 day
