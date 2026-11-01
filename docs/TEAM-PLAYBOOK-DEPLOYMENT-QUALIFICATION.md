# Team playbook deployment and qualification

Owner handoff for Greg, November 1, 2026. Scope: website editor → immutable Team
snapshot → existing GSS app reader. Greg owns the site deployment. Do not change
the [shared route contract](TEAM-PLAYBOOK-API-V1.md) without coordinating with K3.
PPTX/PDF publishing, starter books, new GSS routes and native changes are deferred.

## Current release status — deployed November 1, 21:09 UTC

Greg authorized completing the prerequisites and deployment, and explicitly
accepted breaking the seven historical render jobs' download links. No old
objects were copied or deleted. Saved editor books and account/auth data were
not changed.

- Live revision: `7cf47b1db1c215976aa92a3ffc11da51fabc8132`.
- Pages deployment: `28b9c125-6ea2-465b-ac86-e9e01831614d`, created
  `2026-11-01T21:09:19.562558Z`; production success readback verified.
- [Production deployment](https://github.com/greg643/playbook-generator/actions/runs/41350473428)
  and [CI](https://github.com/greg643/playbook-generator/actions/runs/41350473456) passed.
- Private `JOBS_BUCKET=gss-playbook-jobs`, one-day expiry. Permanent
  `PLAYBOOK_BUCKET=playbook-files` and `AUTH_STATE_BUCKET=gss-playbook-auth-state`
  preserved, along with all production env values and email service binding.
- The renderer now uses `R2_JOBS_*` repository secrets. Effective job-bucket
  write/read/list/delete and **AccessDenied** against both permanent buckets
  passed in [check-only run](https://github.com/greg643/playbook-generator/actions/runs/41350338780)
  and again before the production deployment. Probe objects were removed.
- Preview production bindings/secrets were removed. Callback origin configured.
- Process Playbook workflow was paused for cutover, then re-enabled; active
  state read back after successful production checks.
- 211 Node and 89 Python tests passed locally; Pages Functions compiled.
- Live smoke checks: home 200; anonymous native list/asset and Team discovery
  401; callback without capability 404. These prove route wiring and denial,
  **not** an authenticated Black publication or member/foreign-Team acceptance.
- Durable readback: `output/checks/deploy-live-1101/production-readback.json`.
  Rollback baseline/config fingerprints: `output/checks/deploy-live-1101/before.json`.

**Next:** Greg chooses the existing editor book and intended audience, then run
the live Black qualification below. No live book/snapshot IDs exist from this
deployment. No app, GSS API or TestFlight changes were made.

## Historical preparation status — superseded by the deployment above

**Not ready to deploy until the storage and callback prerequisites below are met.**
Production was checked read-only at 18:33 UTC on November 1:

| Item | Verified state |
| --- | --- |
| Website implementation tested | `e8aefa3d51086ec0417be5f71d334a0f42f060c7` on `codex/team-playbook-k1-k2-1101` |
| Live Cloudflare revision | `aeb0e7e2d8984054a9f8e11161be1068a3488c6a`, deployed October 31 at 00:57:42 UTC |
| GitHub default-branch revision | `dcedd590994c9bcd724579c7ac5d71278b1bd2c4`; the K1/K2 implementation is not yet on remote main |
| Account binding | `PLAYBOOK_BUCKET` → `playbook-files` |
| Private auth binding | `AUTH_STATE_BUCKET` → `gss-playbook-auth-state` |
| Isolated render bucket | **Missing**: no `JOBS_BUCKET` binding and no separate playbook-job bucket in the account inventory |
| Website origin | `PUBLIC_ORIGIN=https://playbook-generator.pages.dev`; Apple enabled |
| Actions callback origin | **Missing**: repository variable `PLAYBOOK_PUBLIC_ORIGIN` is not set |
| Renderer credential scope | **Unverified**: R2 secrets exist but GitHub does not expose their values or Cloudflare token policy |
| Deployment account ID input | Repository secret `CLOUDFLARE_ACCOUNT_ID`, referenced by `deploy.yml`, is not listed; configure it for the intended deployment method |
| Black published book/snapshot IDs | **Not recorded**: no live K1/K2 publication was attempted |

No cloud configuration, credentials, accounts, memberships, source playbooks or
published snapshots were changed during this preparation. No push or deployment
was performed. A missing callback origin or storage isolation is not a reason to
weaken the publisher's checks.

## Local qualification

All tests below ran against website source `e8aefa3`:

| Gate | Result | Evidence |
| --- | --- | --- |
| API, frontend and workflow tests | 210 passed | `node --test tests/api/*.test.mjs tests/frontend/*.test.mjs tests/workflows/*.test.mjs` |
| Python pipeline regression suite | 84 passed | `python -m unittest discover -s tests/pipeline` |
| Actual website builder → actual Swift reader | 11 checks passed | `output/checks/snapshot-native-3aufmt/receipt.json`; app source `13f799de8e2b6fbd8b9e54ec676176ad59f948f8` |
| Editor browser publishing | 7 scenarios passed | `output/checks/deploy-prep-browser-1101/browser-receipt.json` |
| Pages Functions bundle | Compiled successfully with Wrangler 4.131.1 | `output/checks/deploy-prep-functions/` (local build only) |

The browser gate uses synthetic APIs and blocks non-local traffic. It checks
included PNGs, saved revision and printed numbering, Team/default/coaches-only
choices, receipt timing, responsive dialogs, expired authentication, no managed
Teams, unconfirmed publication and edits during save. The native gate compiles
the existing model without modifying the app.

API tests cover first publish, two books/default changes, update failures,
concurrent/stale completions, membership revocation, foreign-Team access,
coaches-only visibility, and denial of old public snapshots after a visibility
change. These results do not establish live identity, bucket permissions or
Black-Team membership.

## Deployment runbook (completed; retain for rollback/redeployment)

1. **Plan a coordinated job cutover.** Stop new render submissions for the
   migration window and let accepted jobs settle. Keep the website account/auth
   data, Apple configuration, `GSS_API_BASE` and `SESSION_SECRET` unchanged. Do
   not apply a whole-bucket expiry rule to `playbook-files`.
2. **Create a private job-only R2 bucket**, for example `gss-playbook-jobs`, and
   bind production Pages `JOBS_BUCKET` to it. Its actual bucket name must differ
   from both account and auth storage; distinct binding names alone prove nothing.
   Configure the job-only bucket's one-day lifecycle. Reconcile existing jobs
   before switching reads: if old downloads must remain available, copy only the
   required `jobs/` data to the new bucket, preserving ownership/status metadata.
   Do not copy credentials, accounts, Team catalogs or auth grants to it, and do
   not delete the original bucket during this release.
3. **Create/verify the renderer's bucket-scoped credential.** It needs object
   read/write access to the job bucket only, with no account, auth or snapshot
   bucket access. Review the Cloudflare token policy and securely update GitHub's
   `R2_JOBS_ACCESS_KEY_ID`, `R2_JOBS_SECRET_ACCESS_KEY`, `R2_JOBS_ENDPOINT` and
   `R2_JOBS_BUCKET` together (mapped to `R2_*` environment variables in Actions).
   Record the scope and date, not the secret values. GitHub secret-name listings
   cannot prove this scope. Do not reuse the old shared-bucket credential.
4. **Set the Actions repository variable**
   `PLAYBOOK_PUBLIC_ORIGIN=https://playbook-generator.pages.dev` (no trailing
   slash/path). The renderer appends `/api/app-publication/complete` and refuses
   redirects. This is a non-secret variable, not a token. Keep it aligned with
   the website `PUBLIC_ORIGIN`.
5. **Verify deployment credentials.** The existing workflow expects
   `CLOUDFLARE_API_TOKEN` and `CLOUDFLARE_ACCOUNT_ID`; account ID is
   `76317c04d814307dbb81c8fe71d287bb`. The deploy credential is separate from the
   renderer credential and must never enter the processing job. Do not share
   production bindings/secrets with a preview host.
6. **Deploy coordinated source.** Integrate the reviewed candidate into the
   GitHub default branch and deploy the matching `dashboard/` revision. Pushing
   dashboard changes to main triggers the existing deploy workflow, so do this
   only when ready. The `repository_dispatch` renderer checks out the remote
   default branch: deploying Pages alone while GitHub still has the old renderer
   does not qualify K1. Verify the callback-aware `process.yml` and
   `pipeline/process_job.py` are on that branch before resuming jobs.
7. **Record the deployed revision from Cloudflare**, not the local branch name.
   Check CI, deploy and renderer workflow results, then sign in with Apple again
   to obtain the new short-lived private GSS publishing grant. Older cookies can
   edit but cannot supply that grant. Run the live sequence below.

Cloudflare documents bucket-scoped Object Read/Write credentials in
[R2 authentication](https://developers.cloudflare.com/r2/api/tokens/).
Do not send credentials in chat, screenshots, issue comments or shell arguments.

## Live Black qualification after deployment

Use an existing editor book chosen by Greg. Record its source ID/revision and
export a local backup before a reversible edit. Select Black from the server's
managed-Team list and record its returned Team ID; never infer an ID from its
name. Use a manager session, a non-manager Black member session (ideally a parent),
and a valid foreign-Team account that is not a Black member. Invalid or expired
tokens are not substitutes for foreign-Team authorization tests. Use secure
authenticated sessions; do not mint or impersonate them.

1. **Publish to the whole Team.** Choose included plays, select Black, enable
   the default star, leave Coaches only off, and publish. Wait for the confirmed
   publication receipt, not just successful PDF rendering. Record job ID,
   published book ID **B**, snapshot **S1** and source revision. Check manager and
   member lists return B with S1, correct counts/order/numbers and default B.
   Download every referenced PNG and each present side's PDF through the asset
   route. Verify byte length and SHA-256 against the manifest; inspect diagrams
   and coach cards against the selected editor plays.
2. **Update the same book.** Make one agreed reversible edit, save and publish
   to Black again. Require the same B and a different snapshot **S2** reflecting
   the new source revision. Confirm the default and visibility are as chosen,
   S2 files verify, and old S1 asset URLs return 404. Do not create a replacement
   source book as a shortcut for this test.
3. **Exercise an accepted publication failure.** Capture S2 and its asset hashes
   first. Start a new publication, then save a reversible source change after
   job acceptance but before completion. Confirm the completion rejects the stale
   revision and does not report a successful receipt. Re-read the Team catalog:
   B must still point to S2 with the same default and bytes, and members must
   still be able to fetch S2. If the job won the race and published first, this
   was not a failure test; repeat with a controlled pause of that test job.
   Do not break shared credentials, revoke real memberships, or corrupt the
   current snapshot to manufacture a failure. Restore the agreed source content.
4. **Check foreign-Team denial.** With the valid non-Black account, request both
   Black's list and a known current B/S2 asset. Both must return 404 without any
   manifest/bytes. With no bearer, require 401. For Black's ordinary member,
   require 200 for the Team-visible list and assets.
5. **Publish Coaches only.** Publish B to Black with the switch enabled and
   record **S3**. Managers must list/read S3; the ordinary member must not see B
   or its hidden default ID, and guessed S3 asset requests must return 404.
   Previously public S2 URLs must also return 404 for that member. If another
   Team-visible book exists, the member's default must resolve to a visible book;
   otherwise require an empty library with null default.
6. **Finish in Greg's intended state.** Confirm whether the final book should be
   Team-visible or coaches-only and restore source text/settings as agreed. If
   this requires a final publication, record its snapshot ID too. A second book
   can exercise default-star switching if Greg selects an additional existing
   book; do not publish private books to parents without that selection.

All reads use the two routes in `TEAM-PLAYBOOK-API-V1.md`; no app or GSS backend
changes are needed. Do not follow redirects with GSS bearer tokens. Record only
status, IDs, safe hashes/counts and workflow references, not request headers or
raw identity responses. Offline/native UI acceptance belongs to K3.

## Completion receipt

Fill this only after the corresponding live operation; local fixture UUIDs are
not production IDs.

| Evidence | Current result |
| --- | --- |
| K1/K2 deployed revision and UTC time | `7cf47b1`, November 1 at 21:09:19 UTC |
| Verified actual job bucket and renderer credential scope | PASS; job roundtrip and two protected-bucket AccessDenied checks |
| Black Team ID and source book ID | Pending authenticated selection |
| Published book B and first snapshot S1 | Pending |
| Successful update snapshot S2 | Pending |
| Failed job ID and unchanged S2/default/hash evidence | Pending |
| Member list/asset and foreign-Team list/asset results | Pending |
| Coaches-only S3 and denial of previous public asset URLs | Pending |
| Final snapshot ID, audience and default choice | Pending |

Deployment prerequisites are green and K1/K2 is live. Resume authenticated live
qualification once Greg selects the book/audience and the required legitimate
sessions are available. Do not describe deployment or anonymous denial checks
as proof of a successful Black-Team publication.
