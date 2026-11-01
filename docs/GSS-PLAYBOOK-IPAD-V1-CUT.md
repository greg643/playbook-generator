# GSS playbooks in the app: first release cut

Addendum to `GSS-PLAYBOOK-IPAD-INTEGRATION.md` · Greg via Fable, 2026-10-31.
The integration doc stands as the full design. This page is what to build first and what the screens are.

## Product rules (Greg, 10/31)

- A coach can publish several playbooks and see all of them in the app. One is the **default** and opens directly.
- A book opens as an overview of all its plays, not a single play. The coach picks a play from the overview and it fills the screen. Swiping between plays is secondary.
- A book is published **to a Team**. Today only Greg has plays in the editor, so the book is not the coach's private library in the app; it is the Team's playbook. At publish time the coach picks the GSS Team (one where their account is a manager). Every member of that Team, parents included, can open it; members of other Teams cannot. A per-book switch "Coaches only" exists for a book the coach does not want parents to see; default is the whole Team.
- Works on iPhone and iPad. iPhone shows the overview in two columns and the full-screen play is usable; the iPad is the real sideline surface.
- No new tab. The tab bar is full (Home, My Teams, Scoring, Tag Reader, Menu). Playbook is a row in the Team hub, between Players and Highlights, with a book icon. When the Team has no published book the row is absent for members and reads "Publish a playbook from the web editor" for managers.
- Overview has two forms: a **grid** (4x4 for 16 plays, the same layout as the printed coach card, so the sideline card and the iPad match) and a **list** by number and name. The coach's last choice is remembered.

## Screens

1. **Playbook row in the Team hub.** Tapping it opens the Team's default book straight to its overview. If the Team has more than one published book, a book picker sits at the top of the overview (name, 5v5/6v6, play count, published date); the manager's star sets the Team default and that choice is published with the book, so every member opens the same default. "Download for offline" and "Ready offline" per book, as in the doc.
2. **Book overview**. Offense / Defense switch at the top. Grid or list toggle. Grid tiles show the play image with its printed number and name; 4x4 on iPad, 2 columns on iPhone. List rows show number, name and a small thumbnail. The order is the published order and is never renumbered.
3. **Play**. Full screen, pinch zoom, number and name stay visible. Previous / Next and swipe move to the neighbouring play in the same side. Back returns to the overview at the same scroll position. Switching to replay and back keeps the open play.
4. **Coach card** button on the overview opens the published PDF for that side.

## Build order

1. Website: **Make available in GSS app** writes a snapshot (manifest, play PNGs, offense and defense coach PDFs) under the coach's account. Only included plays. One rule: a failed publish never replaces the current snapshot.
2. Playbook service: two read routes, list books for a Team and fetch asset. The app sends its GSS session and the Team ID; Playbook validates the session against the existing GSS identity route (the same backend it already uses for Apple sign-in), checks that the account is a member of that Team through the existing GSS team-membership read, and caches both briefly. Publishing to a Team requires the manager role on that Team. No new GSS backend routes in v1.
3. App: screens 1–4 above, offline download under the account key, logout clears account-local downloads. Simulator XCUITests: overview grid and list, tap to play and back, default book opens first, logout during download.
4. Starter playbooks, if few coaches have drawn plays yet. Greg answers 5v5 or 6v6 and whether "two default playbooks" means one per format.

## Deferred to the second release

Signed service assertions and GSS-side proxy routes, snapshot garbage collection and retention, ETag/304, rate limits, wristbands in the app, sharing, clip association, editing.

## Acceptance

- Two published books: the default opens directly; the other is one tap away; changing the star changes which opens next time.
- A 16-play offense book shows as a 4x4 grid matching the printed coach card, and as a list; tapping any tile opens that play; Back returns to the same place.
- Offline: airplane mode, kill the app, reopen, the default book opens and every play renders.
- A parent on the Team opens the Team's playbook; a parent on another Team cannot list or fetch it; a "Coaches only" book is invisible to parents.

## Bring your own playbook (Greg, 11/1)

Coaches who already have a playbook should not redraw it. Two more sources feed the same snapshot and the same grid, list and tap-to-zoom viewer:

- **PowerPoint (K5).** A completed PPTX Import can be published to a Team exactly like an editor book. Entries come from the full slides in slide order, named from the slide title or "Play n". Greg's book is the first acceptance target.
- **PDF (K6).** New upload beside PPTX Import. One tile per page. A page that holds a grid of plays stays one tile and the coach zooms into it; names from the page's text layer when present, else "Page n", editable before publishing. Splitting a grid page into separate tiles by drawing boxes is a later option if coaches ask.

The manifest marks each entry's `source` as `editor`, `pptx` or `pdf`. The app shows "Imported" for the last two and never offers editing. The same viewer serves all sources.

Upload and conversion stay on the website/Actions runner for v1 (LibreOffice and Poppler). Phone share-sheet upload is a later step after K6 and must use the same pipeline. Site deployment remains Greg's.
