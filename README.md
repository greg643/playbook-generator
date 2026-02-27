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
