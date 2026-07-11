# Review of hardening changeset (vs baseline 016f3d0) — 2026-07-11

Reviewed by Claude (xhigh-effort pass: 10 finder angles → manual verification against
working tree + baseline + Greg's real deck + test runs). Audience: Codex 5.6 and future
sessions. Working tree at review time: uncommitted hardening changes, tree unchanged
since snapshot `review-scope-2026-07-11.diff` (in Claude's memory dir).

Overall: the hardening direction is right and most of it is well built (auth versioning,
conditional R2 writes, quota system, input caps, CI, tests). But two changes to the PPTX
pipeline **break Greg's actual deck**, and several edge races deserve a pass before real
use. Test suites pass (node 31/31, pytest 18/18) — note the fixtures encode the *new*
behavior, so green tests do not clear the blocking items below.

## BLOCKING — the primary real-world input fails

> **STATUS UPDATE (same day, by Claude):** Findings **1**, **2**, and **13** are
> FIXED in the working tree. Section separators are now a case-insensitive free-text
> match ("OFFENSE"/"DEFENSE" anywhere in the slide text) on slides *without* a field
> rectangle (so play slides mentioning a section stay plays; ambiguous both-word
> slides are not separators). Single-section decks now produce the outputs they can;
> zero-producible raises a ValueError carrying the new `SECTION_HINT` guideline, which
> also appears on the upload page (index.html `.format-hint`) and in help.html.
> `produced` verification is scoped to `OUTPUT_FILENAMES`. Tests updated/added:
> pytest 23/23, node 31/31 pass; verified end-to-end against the real deck
> (`Flag_Playbook_6v6_2026.pptx` → 16 offense + 5 defense plays detected).
> Findings 3-12 and the cleanup section remain open for Codex.

### 1. Exact-match section headers reject the real deck (zero plays)
`pipeline/playbook_pipeline.py:76` — detection changed from baseline's substring match
(`"OFFENSE" in all_text`, gated to sparse slides) to exact whole-shape-text equality
(`"OFFENSE" in text_values`). **Verified against Greg's actual 2026 deck**
(`~/projects/playbook_generator_pptx/Flag_Playbook_6v6_2026.pptx`): its header slides
say **"6v6 OFFENSE"** (slide 1) and **"6v6 DEFENSE"** (slide 18) as single text shapes.
`"6V6 OFFENSE" != "OFFENSE"` → `current_section` stays `None` → every slide skipped →
zero plays → job fails. The Fall 2025 deck likely has the same style.
- Fix idea: keep the good part of the change (repeated sections continue numbering; no
  shape-count gate) but match per-shape text with a word-boundary test, e.g.
  `re.search(r"\bOFFENSE\b", t)` on short header-ish texts (e.g. len ≤ 30 or ≤ 3 words),
  choosing OFFENSE/DEFENSE by which word appears. Beware the converse: don't let a play
  slide with a stray exact "OFFENSE" label get swallowed as a header — prefer requiring
  the text to be short/dominant on the slide.
- Tests: `tests/pipeline/test_analyze_playbook.py` only uses exact `OFFENSE`/`DEFENSE`
  fixtures. Add a fixture with "6v6 OFFENSE"-style decorated headers (real-deck shape).

### 2. Single-section deck + default options → whole job fails
`pipeline/playbook_pipeline.py:717-739` — `generate_all` skips a section with no images
(`if gen_defense and defense_images:`) but then raises `RuntimeError("Generated output
set did not match the request …")` because expected ≠ produced. Upload path defaults all
four outputs on (`dashboard/functions/api/upload.js:109-112`, absent → `true`), and the
uploader cannot know the deck's sections in advance. Baseline delivered the available
section's PDFs. The images-mode endpoint has a proper guard
(`dashboard/functions/api/generate.js:157`) — the PPTX path has no equivalent, and
`process_job.py`'s message mapping (~line 319+) has no branch for this text, so the user
sees the generic "Processing failed unexpectedly."
- Fix idea: in the PPTX path, reconcile expected vs available *after* `analyze_playbook`
  — drop selected-but-empty sections (record a warning in status.json, e.g.
  `"warnings": ["No DEFENSE section found — defense outputs skipped"]`), and only fail
  if *nothing* can be produced (with a friendly ValueError). Keep the strict
  produced-vs-expected check against the reconciled set — it's a good invariant.

## HIGH

### 3. Stale-timeout counts Actions queue time; processor overwrites terminal status
`dashboard/functions/api/status/[[jobId]].js:41-46` measures the 15-min timeout from
`createdAt`, but `timeout-minutes: 10` in process.yml bounds only *run* time — GitHub
queue delay is unbounded. A queued-then-successful job gets marked
`error: "Processing timed out"` mid-poll. The conditional write + re-read on loss
(lines 54-68) is well done, but `process_job.update_status` (pipeline/process_job.py:85)
is an **unconditional put** — the still-running worker later overwrites the terminal
error with steps and `complete`, after the UI already showed failure and released the
quota slot (`finishJobSlot`, line 72), letting a duplicate run start.
- Fix idea: (a) have `process_job` heartbeat (`updatedAt` on every status write) and
  measure staleness from the last write, not `createdAt`; (b) make `update_status`
  respect terminal statuses the way `mark_job_failed.py` does (read-check-write, ideally
  conditional) so a worker never resurrects a timed-out/cancelled job.

### 4. `PUT /api/plays` 403s clients loaded before the deploy (autosave dies)
`dashboard/functions/api/plays.js:88` rejects any doc where `doc.ownerId !==
user.userId`. Baseline `editor.html` contains **zero** occurrences of `ownerId`
(verified), so every tab open across the deploy autosaves with `ownerId === undefined`
→ 403 forever until manual reload; edits in that session are never persisted. The
binding itself is a good idea (new client sends it, editor.html:2175).
- Fix idea: transition-accept missing `ownerId` for a release (log it), or return a
  distinct error the old client turns into a "reload required" prompt. Given the small
  user base, at minimum deploy at a quiet hour and note it in the release steps.

## MEDIUM

### 5. `getUser` catch-all turns transient R2 errors into sign-outs
`dashboard/functions/_lib/auth.js:272` — the whole verification (including the new
account fetch at line 250) is in one `try { … } catch { return null }`. A transient
storage exception → `null` → 401 → the frontend calls `transitionToSignedOut()` and
aborts in-progress UI, though the cookie is valid. Baseline had no R2 read here, so
this failure mode is new.
- Fix idea: catch only parse/shape errors; let infrastructure exceptions propagate to
  the endpoint's 500 (clients treat 5xx as retryable, and both HTML pages already do).

### 6. Quota slot never released across UTC midnight
`dashboard/functions/_lib/jobs.js:108-117` — `updateReservation` builds the key from
*today* only; a job reserved at 23:55 UTC finishing 00:05 isn't found (`findIndex ===
-1` → silent return), so the slot stays active until `activeUntil` (~30 min). UTC
midnight = 7-8 PM ET, prime coaching hours; with maxActive=2 this can 429 real users.
`listUserJobIds` (same file, line 149) already reads a two-day window — do the same here.

### 7. `mark_job_failed` can clobber a completed job on a transient read error
`pipeline/mark_job_failed.py:52` — `except Exception: current = {}` swallows *any*
status-read failure, then the terminal-status guard passes vacuously and the finalizer
overwrites `complete` (files list lost → PDFs undownloadable). Only treat
missing-key as empty; retry/re-raise other errors (the step is `continue-on-error`).

### 8. `merge_block_caps` drifted from the inline version it replaced
`pipeline/play_geometry.py:17-30` vs baseline `pptx_to_editor.py:254-283` — the helper
added requirements the original didn't have: cap `end=='none'` + not dashed + length ≥
0.005, route `end=='none'` (was: anything except `'ball'`), not dashed, and **same
color as the cap**. A black cap on a colored route stub (or a dashed motion route) no
longer folds into `end='block'` — imports render a stray tick line + missing block
terminal. Some tightening is defensible (not overwriting arrowheads); the color-equality
check is the most likely real-world regression. Verify by importing a real deck with
block T's; consider dropping the color requirement.

### 9. Field-rectangle fallback removal makes play loss *silent*
`pipeline/playbook_pipeline.py:93-99` — baseline fell back to the largest shape when no
"rectangle"-named shape existed; now such slides are skipped with only an Actions-log
print. Removal is deliberate (comment: logos/photos as crops) and reasonable, but a
deck whose fields are pictures/freeforms now loses plays **silently** — job completes,
PDFs missing plays. Surface it: count skipped play-candidate slides into a
`warnings` field in status.json (pairs with the warning channel from finding 2).

## LOW

10. **process.yml:29** — the detect-mode step lost baseline's `|| echo pptx` fallback;
    one transient R2/network blip now fails the whole job (baseline recovered; the
    worker itself retries reads 3×). Add a retry loop or fallback. (The new `assert
    mode in {...}` validation is good — keep it.)
11. **login.js:100** — after losing the legacy-hash upgrade race, the re-read record is
    passed to `hashPassword(password, latest.salt, latest.iterations)` *before* its
    shape is validated; a tombstone/malformed record throws → 500 instead of 401.
    Reorder: validate `latest.userId/disabledAt/salt/iterations` first.
12. **process_job.py:62-71** — `scrub_job_payload`'s `while True` never inspects
    `delete_objects`' per-key `Errors` (Quiet mode); a persistently undeletable key
    spins until the Actions timeout. Bound iterations or check Errors.
13. **playbook_pipeline.py:730** — `produced` globs *all* `*.pdf` in output_dir, so a
    CLI run into a folder containing unrelated PDFs (e.g. `notes.pdf`) raises
    "unexpected: notes.pdf" despite full success. Scope the check to
    `OUTPUT_FILENAMES.values()`.

## Cleanup / consolidation notes (non-blocking, worth a tidy pass)

- **`dashboard/_headers` does not apply to Pages *Functions* responses** — the `/api/*
  Cache-Control: no-store` block is a no-op (handlers already use `jsonNoStore`, which
  covers JSON, but raw responses like the PDF download rely on hand-set headers). If
  API-wide headers are wanted, add `dashboard/functions/_middleware.js`.
- Duplication introduced by the changeset (each a drift risk; consolidate in `_lib`):
  `contentLengthTooLarge` in generate.js/upload.js duplicates `requestBodyTooLarge`
  (auth.js); `quotaResponse`/`isFile`/cleanup-wrapper duplicated between generate.js and
  upload.js; sessionVersion clamp inlined in login.js ×2 + recover.js while
  `accountSessionVersion()` exists unexported; `staleAfterMs` re-implements
  `boundedEnvInt` (jobs.js); the ~100-line generation/polling state machine is
  near-verbatim in index.html and editor.html (CSP `script-src 'self'` allows a shared
  static js); the bounded-JSON-body intake sequence is copy-pasted across 5 endpoints
  (`readBoundedJson` helper); `mark_job_failed.py` re-implements `get_r2_client`/
  `job_is_cancelled`/`update_status` from process_job (which it already imports);
  `accounts/{userId}/playbook.json` is hand-built ×4 in delete-account.js while plays.js
  has `playbookKey()` — a data-retention hazard if the key ever changes; play-image
  budget (22 images, name grammar) declared in process_job.py, regexes in
  playbook_pipeline, and generate.js instead of once in input_safety.py.
- Efficiency (hot paths): status endpoint does 4 serial R2 reads per 3-second poll
  (Promise.all the owner+status reads); `finishJobSlot` conditional-PUTs even when the
  slot is already finished (early-return; polls of finished jobs currently write
  forever); `cleanupFailedJob` re-GETs up to 24 keys to verify deletes R2 already
  guarantees; plays.js PUT now GETs the full stored doc even on the legacy path where
  only the ETag is used (`head()` suffices); delete-account runs `ensureDeletionRecord`
  twice and deletes playbook.json twice outside the resweep; process_job's upload loop
  calls `ensure_job_active` twice per iteration back-to-back.
- upload.js:95 — the `textBytes > 100` check is dead code (the allowlist above caps
  text fields at ~20 bytes); delete it.
- README local-dev section instructs `python3.11 -m venv .venv` — on Greg's machine
  that conflicts with his global convention (shared `~/.venv`, Homebrew interpreter).
  Reword to "any Python ≥3.11 environment with `pip install -r requirements.txt`"
  (CI is unaffected).
- README documents auth rate-limiting as a manual Cloudflare dashboard step; each guess
  costs a 600k-iteration PBKDF2 on the Worker. Fine for launch at this scale, but a
  code-level per-account/IP throttle (same conditional-put pattern as reserveJobSlot)
  would make the edge rule defense-in-depth rather than the only layer.
