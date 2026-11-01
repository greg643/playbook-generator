# Team Playbook API v1

Canonical wire contract for K1/K2 and the flag-app K3 reader. The product cut is
[GSS-PLAYBOOK-IPAD-V1-CUT.md](GSS-PLAYBOOK-IPAD-V1-CUT.md); it overrides the older
integration proposal. In particular there are **no new GSS backend routes** and
no `/v2/playbooks` proxy in v1. Site deployment belongs to Greg.

## Native read routes

Origin: `https://playbook-generator.pages.dev`. Every request sends
`Authorization: Bearer <current GSS product session>`. Do not send the Playbook
cookie, an Apple identity token, an account ID, or an R2 credential instead.
Redirects must not be followed with this bearer.

| Method/path | Success |
| --- | --- |
| `GET /api/app/playbooks/{teamId}` | Library JSON below, including complete manifests |
| `GET /api/app/playbooks/{teamId}/{bookId}/{snapshotId}/{assetId}` | Verified PNG or PDF bytes from the current authorized snapshot |

Team IDs are `T-` plus a lowercase v4 UUID; published book/snapshot IDs are
lowercase v4 UUIDs. Asset IDs contain 1–128 ASCII letters, digits, `_` or `-`.
An editable source book ID is not the published book ID. Neither storage keys
nor media URLs are accepted from the client or returned in manifests.

For each request the service calls existing authenticated `POST /v2/me`
(`include_teams: false`) and `POST /v2/teams/get` (`team_id`). It requires an
enabled, migration-complete account and an active Team with role `leader` or
`member`. A manager is a `leader` who is not a family member. Identity and
membership are currently checked fresh, without a cross-request token cache.
Parents need no separate editor account. Only managers see coaches-only books.

Responses use `Cache-Control: private, no-store`. Asset responses include exact
Content-Type, Content-Length and `X-Content-Type-Options: nosniff`.

| Status | Meaning / client action |
| --- | --- |
| 200 | Library or verified asset; an authorized empty Team returns an empty library |
| 401 | Missing/rejected session; use the existing GSS refresh flow once, then sign in |
| 403 | GSS account is not enabled/migration-complete; deny this resource, not a generic global session revocation |
| 404 | Invalid, foreign-Team, hidden, deleted, or superseded resource; do not reveal existence |
| 409 | Invalid/conflicting stored publication; refresh, never treat as an empty success |
| 503 | Upstream/storage/verification unavailable; preserve a valid older offline download |

Error body is `{ "error": "safe human-readable message" }`. Raw upstream text,
tokens, private drawing content and storage paths must not be logged/reflected.

## Library and immutable manifest

Library fields:

| Field | Type / rule |
| --- | --- |
| `schema` | `gss.team-playbook-library.v1` |
| `teamId` | Exact requested Team ID |
| `canPublish` | Boolean, verified manager authority |
| `defaultBookId` | Visible book ID or null when empty |
| `books` | 0–100 complete snapshot manifests |

The Team default opens directly. If it is coaches-only and the reader is a
parent, the first visible book is that reader's default; a hidden ID is never
returned. Updating the default and updating the snapshot are one atomic catalog
change, not a client-local preference.

Snapshot fields (all required):

| Field | Type / rule |
| --- | --- |
| `schema` | `gss.team-playbook-snapshot.v1` |
| `bookId`, `snapshotId`, `teamId` | Published identities; stable book, immutable snapshot |
| `name` | Normalized saved book name, at most 80 characters |
| `sourceRevision` | Exact saved source revision UUID |
| `publishedAt` | UTC ISO-8601 with fractional seconds |
| `coachesOnly` | Boolean; default false at publish |
| `format` | `5v5`, `6v6`, or `mixed`, derived from selected plays |
| `plays` | Ordered entries, offense then defense, original source order within each side |
| `assets` | Exactly the referenced PNGs and side PDFs; no extra files |
| `coachCards` | Object keyed by each nonempty side (`offense`, `defense`) |

Each play has `id`, `sourcePlayId` (same stable source ID), `side`, `name`,
`number` (original integer 1–999), `displayNumber`, `playersPerSide` (5 or 6),
and `imageAssetId` (`image-` plus the play ID). Defense numbers 1–26 display as
A–Z; otherwise the display label is the decimal number. Never renumber selected
plays. Only `included !== false` plays are published. IDs are unique across the
snapshot, numbers unique per side. Limits: 64 offense + 24 defense plays.

Each asset has `id`, `mediaType` (`image/png` or `application/pdf`), `bytes`
(positive integer), and `sha256` (64 lowercase hex digits). PNGs additionally
have integer `width` and `height` (1–8192, product at most 16,777,216). Limits are
4 MiB/PNG, 16 MiB/PDF, 64 MiB/snapshot, 90 assets. Each side's coach card is
`{ "assetId": "coach-offense|coach-defense", "playsPerPage": 1|2|4|6|9|16 }`.
Missing sides have no PDF entry. Each stored manifest is bounded to 128 KiB.

Native decoding rejects unsupported schemas, duplicate/bad identities, extra or
missing referenced assets, inconsistent formats, dimensions and limits. Native
download verifies byte count, SHA-256 and decoded PNG/PDF before marking ready.
The service verifies stored object size/checksum before streaming it.

## K1 publication lifecycle

The editor saves the book first, then uses its existing same-origin cookie
`POST /api/generate` multipart path with the included PNGs and an
`options.appPublication` selection: `teamId`, `sourceBookId`, `sourceRevision`,
`coachesOnly`, `makeDefault`, and optional side `layouts`. The server assigns the
job/operation ID. The browser does not choose permanent keys or provide a GSS
token in the form. Its Apple-authenticated session has a short-lived encrypted
GSS grant, bound to that exact cookie, account and origin.

Acceptance rechecks both account identity and Team manager authority, reads the
owned saved revision, validates selection, and persists an ordered intent.
The isolated renderer receives job files and a one-job callback capability, not
the GSS session or account/auth bucket credentials. It completes through
`POST /api/app-publication/complete`, even after the browser closes. Callback
configuration is validated before rendering; retries retain the same identity.

Completion rechecks authority, source revision, ownership and cancellation,
copies verified files to private immutable snapshot keys, writes the manifest,
then conditionally changes the Team catalog pointer/default. A partial copy,
failed render/dispatch, lost authority, source edit or stale completion never
replaces the previous complete book. Retrying an acknowledged operation is
idempotent. An unconfirmed callback is not presented as successful publishing.
Changing a book to coaches-only prevents fetching its previous public snapshot.

## K3 integration and offline boundary

flag-app uses this contract in `Sources/Teams/TeamPlaybook.swift` and
`TeamPlaybookClient.swift`; there is no separately defined transport contract.
The UI is in the Team hub, not a new tab. Account/session-generation checks fence
late async writes. A complete verified download is committed atomically; a
replacement failure keeps the older complete copy. Logout purges local content.
Reconnect reconciles Team access and removes revoked content. An offline device
cannot learn of a new remote revocation until reconnect/sign-out.

## Verification and deployment ownership

Behavior tests: `tests/api/app_publications.test.mjs` (authorization, two books,
star, failed/stale/concurrent publications, revocation and private bytes),
`app_snapshots.test.mjs` (selection/schema), `app_web_session.test.mjs`
(session boundary), and `tests/pipeline/test_app_publication.py` (completion).
The browser gate is `tests/frontend/editor_publication_browser.mjs`.
K3 uses native unit and Simulator UI tests; those are not live publishing proof.

Before Greg deploys: follow README's storage migration order, bind a distinct
JOBS_BUCKET, restrict Actions R2 credentials to that bucket, set
`PLAYBOOK_PUBLIC_ORIGIN=https://playbook-generator.pages.dev` as a GitHub
repository variable, and supply the deployment account ID. Preserve existing
account/auth data and SESSION_SECRET. Test a real manager publish plus member,
foreign-Team and coaches-only reads after deployment. No push/deploy or new GSS
backend routes are authorized by this document.
