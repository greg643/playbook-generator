# GSS playbooks on iPad

Design handoff for the flag-app project · October 31, 2026

## First release

Coaches create their plays on the GSS Playbook Editor website, then open them in
the GSS app using the same Apple Account. On an iPad they can swipe through plays,
open a coach-card PDF, and return to replay without changing apps. Downloaded
playbooks work without field Wi-Fi.

Keep editing and printing setup on the website initially. The app is a read-only
viewer. This document proposes the integration; the viewer and its APIs are not
implemented by the Apple-only website release.

Start with saved editor playbooks. PowerPoint (PPTX) Import currently produces
conversion jobs, not persistent editor books. A later **Save to app library**
action can publish a completed job's rendered play images and PDFs through the
same snapshot pipeline, without pretending those images are editable diagrams.
Verify job ownership and copy its assets before expiry; never publish from an
arbitrary job ID or URL supplied by the client.

## Coach experience

1. On the website, select a named playbook and check the plays to include, just as
   for a coach card. Choose **Make available in GSS app**. Confirm the selected
   offense/defense counts and publish a saved version.
2. In the app, open **Playbooks** alongside the existing replay experience. Show
   the coach's published books, with name, format, play count and update date.
3. Choose **Download for offline use**. Once every required file is verified,
   show **Ready offline**. Opening a book displays one play at a time with large
   Previous/Next controls, swipe navigation, an offense/defense switch and a
   thumbnail picker. Keep the printed play number and name visible.
4. Offer **Coach card** to open the published PDF. Leave wristband printing as
   a website action in the first release; it need not clutter the sideline viewer.
5. When the website publishes changes, show **Update available**. Retain the
   complete previous download until its replacement is ready. Display which
   version is open, especially while offline.

Use a full-screen viewer that preserves diagram proportions in portrait and
landscape and supports zoom. Switching to replay and back must retain the open
play. A side-by-side replay/play view can follow later if it does not compromise
playback performance or reduce play diagrams to an unreadable size.

An empty library says **Create your first playbook on the web** with a website
link. An unpublished website book is not an app error: explain that the coach
must make it available from the editor. New app users should not need a second
password, and creating an app account must not silently create or overwrite books.

## Existing foundations

| Area | Current behavior | Integration requirement |
| --- | --- | --- |
| Identity | Website Apple sign-in uses the GSS verifier and its `account_key` | Resolve the same verified identity server-side |
| Sessions | The website has its own signed cookie; the app has a GSS product session | Do not exchange or share these session secrets |
| Storage | Private R2 records are owned by an immutable local Playbook user ID | Map the GSS identity to that owner before each read |
| Books | Up to 20 named books; season is currently part of the name | Preserve names and IDs; explicit season metadata can follow |
| Plays | Schema 2 supports 5v5/6v6, offense/defense, numbers and included/benched state | Preserve ordering and selection; do not renumber on the iPad |
| Printing | Coach layouts support 1, 2, 4, 6, 9 or 16 plays, independently per side | Reuse generated PDFs and existing print geometry |
| App identity lifecycle | `AccountStore` and request identity include an account key and session generation | Scope every request and cache update to both |

The same Apple login establishes identity, not automatic data sync. Apple email
addresses, including private relay addresses, are not account-linking keys. An
Apple-linked website account may still store its canonical record under a legacy
email key; the existing Apple lookup must follow that index without moving data.

Relevant code to inspect when implementing:

- Playbook: `dashboard/functions/_lib/apple.js`, `_lib/auth.js`,
  `_lib/playbooks.js`, `api/playbooks.js`, `api/plays.js`, `api/generate.js`,
  and `dashboard/editor.html`.
- GSS app repository: `AUTH_AND_DATA_MODEL.md`, `Sources/App/Account.swift`,
  `Sources/Upload/PresignerClient.swift`, and the backend `/v2/auth/apple` and
  `/v2/me` handlers. Preserve its enabled-account and migration-complete gates.
- Authentication setup: [Sign in with Apple](apple-sign-in.md).

## Authentication and API boundary

Use the existing GSS backend as the app-facing API. The app sends its normal GSS
product bearer token. That backend verifies the session and account status, then
calls a narrowly scoped Playbook read service. The browser continues using its
existing cookie-based APIs; do not make those accept unverified app account IDs.

```text
iPad app -- GSS product session --> GSS backend
                                    |
                           signed read-only request
                                    |
                                    v
                              Playbook service
                                    |
                         verified GSS account key
                                    |
                                    v
                         local owner -> published books
```

For the server-to-server bridge, use a short-lived signed service assertion with
an explicit issuer, Playbook-service audience, verified GSS account key, read-only
scope, expiry and key ID. Prefer asymmetric signing so Playbook holds only a
verification key. Start with a maximum 60-second lifetime and bounded clock skew;
rotate keys with an overlap period. The GSS backend derives the account key from
its session, never from a client request parameter. Restrict the bridge to the
read routes and verify algorithm, signature, issuer, audience, scope and time.
Never log assertions or return them to the app.

The Playbook service resolves the Apple index, checks the canonical account is
active, and verifies that each requested book, snapshot and asset belongs to it.
Return not-found for inaccessible resource IDs. No matching website account
means an empty library with a setup link, not an implicit merge. A deleted
Playbook account cannot be revived merely by asking to list its books.

Initially proxy bounded asset downloads through these authenticated services.
This avoids public R2 URLs and reusable bearer links. If bandwidth later warrants
signed URLs, scope them to one asset and a short expiry, exclude them from logs,
and accept that already-issued links remain usable until they expire. Browser
cookies, Apple ID tokens, service signing keys and R2 credentials must never
become app download credentials.

Proposed app-facing routes, subject to the flag-app backend's route conventions:

| Route | Purpose |
| --- | --- |
| `GET /v2/playbooks` | List published books available to the authenticated coach |
| `GET /v2/playbooks/{bookId}/manifest` | Get the current immutable snapshot manifest; support ETag/304 |
| `GET /v2/playbooks/{bookId}/snapshots/{snapshotId}/assets/{assetId}` | Download an owned image or PDF with bounded size and explicit content type |

No endpoint accepts a client-supplied owner ID. Limit list size, manifest size,
asset count, image dimensions and download concurrency. Rate-limit by verified
account. Use private/no-store caching on authenticated responses unless a reviewed
private-cache policy is added; ETags still support explicit app refreshes.

## Published snapshots

Publish a versioned, immutable manifest plus image/PDF assets. This is separate
from the editable JSON and from temporary generation jobs. Do not store a job
download URL as a permanent app library entry: jobs can expire or be scrubbed.

Before expanding production access, complete the existing account/job bucket
isolation work described in the README. The October account inspection still
found a shared account/job bucket. Generation runners must have access only to
job data, not credentials, saved books or the permanent app snapshot library.

The website first saves the selected book, then submits its source revision and
print settings to a new authenticated publish operation. The server verifies
ownership, selection and revision. It validates rendered images using the same
bounded image-upload rules as current generation. App publishing must not depend
on the coach leaving the browser open after a job has been accepted.

Use a distinct private snapshot prefix or bucket with explicit retention. Copy
successful assets into it, verify completeness and hashes, then conditionally
update the book's current-snapshot pointer. A failed or outdated publish leaves
the previous snapshot current. Duplicate requests use an idempotency key; a late
job cannot overwrite a newer publication. Persist the source revision and
selection used so that editing during generation cannot produce a mixed version.

The manifest should contain:

- `schemaVersion`, opaque `bookId`, `snapshotId`, `sourceRevision`, `publishedAt`,
  book `name`, and a declared format (`5v5`, `6v6` or `mixed`). Do not assume every
  existing book is uniform merely because it has a default player count.
- Ordered play entries: stable source play ID where available, snapshot entry ID,
  `name`, printed `number`, `side`, `playersPerSide`, and image asset ID. Snapshot
  entry IDs must not depend on mutable names or printed numbers.
- Asset descriptors: opaque ID, media type, byte length, SHA-256 and image
  dimensions. Avoid absolute storage keys or external URLs in the manifest.
- Optional offense and defense coach PDF descriptors and their chosen layouts.
  Omit a missing side cleanly. Odd counts retain blank grid slots, as printing does.

Publish only included plays by default, with a clear confirmation of the selection.
Benched plays remain private to the editable book unless a future explicit option
includes them. Empty selections cannot be published. Labels and drawing content
are data, not HTML to execute. For the first viewer use raster images rather than
reimplementing the editor's drawing engine in Swift.

Unpublishing removes a book from the online library and blocks future downloads.
Deleting a book or Playbook account must also retire its snapshots and schedule
asset cleanup. Maintain ownership until cleanup finishes so dangling assets cannot
be served. Define a retention period for superseded snapshots before rollout;
garbage collection must not delete a current or in-progress publication.

## Offline storage and account changes

Store downloads under the verified GSS account key, book ID and snapshot ID in
app-private storage with iOS file protection. Exclude reproducible downloads from
device backups and do not expose a shared public cache. Keep partial downloads
separate, verify lengths/hashes and decode limits, then atomically mark the
snapshot ready. Persist manifests so airplane-mode launches can open completed
books without fetching an expired URL.

Capture `V2AuthenticatedRequestIdentity` for every load and download. Before
applying a response or committing a file, confirm both `accountKey` and
`identitySessionGeneration` still match. Cancel tasks, clear views and remove
account-local downloads on logout or account switching; a late response must not
repopulate them. Do not create a second identity manager for this feature.

An offline device cannot learn immediately about revocation or deletion. State
this limitation plainly: completed downloads remain viewable until the app next
validates access, the user signs out, or removes the download. On reconnect,
reconcile the library and remove revoked content before fetching new assets.
Preserve existing GSS offline-access policy; do not bypass a mandatory account
gate to enable this feature. Shared device/logout behavior is a release test.

## Starter playbooks and future sharing

Starter books are a separate follow-up, not required for the viewer. Use a private,
versioned template catalog containing approved diagrams with clear rights to
redistribute. A supplied NFL example is not automatically licensed for inclusion
in every account; use original or permission-cleared artwork and naming.

Suggested setup: ask **5v5 or 6v6?** and offer the corresponding starter set. The
owner must decide whether “two default playbooks” means one of each format or two
books for the selected format before implementation. Store that choice in config.

Provision by copying templates into independently owned books with new IDs, using
an idempotent account/template-version marker. Respect the 20-book cap. Never
replace existing work, recreate a deliberately deleted starter on every login,
or push later template edits into a coach's customized copy. Existing coaches
get an explicit **Add starter playbooks** action.

Keep the current JSON export/import sharing path. Team-library sharing comes
later and needs explicit owner/editor/viewer roles, invitations, revocation,
copy-versus-live-library semantics and GSS team authorization. Knowing a book ID,
sharing an Apple login or being on a roster is not permission to read a playbook.

## Delivery plan and acceptance

1. Confirm GSS web/native identities match for a test coach and audit the existing
   account gates. Choose the app's Playbooks entry point, retention policy and
   rollout flag. Keep live private playbooks out of fixtures.
2. Build website publishing, persistent snapshots and authenticated bridge reads.
   Test owner isolation, stale publication rejection and account deletion cleanup.
3. Build the app library, image viewer, PDF view and atomic offline downloads.
   Exercise portrait/landscape, large text, VoiceOver and account-switch races.
4. Pilot behind a feature flag with a few coaches before league-wide release.
   Measure publish/download failures and timing without logging play content,
   account emails, credentials or download tokens.

Release acceptance:

- The same coach sees the intended published book through web and native identity;
  a different coach cannot list or fetch it by guessing any ID.
- 5v5 and 6v6, offense-only, defense-only and mixed-side selections preserve
  diagrams, order, printed numbers and PDF layouts. Benched plays are excluded.
- First publication failure shows a useful retry state; update failure retains
  the previous complete book. Unsupported manifest versions fail safely.
- A full download survives an airplane-mode app restart. Partial or corrupt
  downloads are not labeled ready. Storage-full and interrupted-download paths
  recover without discarding a valid older version.
- Logout/account switch during a slow download never exposes one coach's content
  to the next. Deleted/disabled accounts and unpublished books cannot download;
  offline revocation behavior matches the stated policy.
- Opening plays and returning to replay preserves each screen's state and does
  not interrupt recording/playback unexpectedly. Test on a real supported iPad.

Not in the first release: native play editing, two-way JSON sync, live team
collaboration, automatic association of plays with replay clips, mandatory GSS
website-domain migration, or automatic starter-book rollout. These can build on
the same identity and snapshot boundary without delaying a useful sideline viewer.
