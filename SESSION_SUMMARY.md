# ShareUp — Session Summary

Covers: SharedView consolidation, Drive slip storage migration, Safari session-persistence investigation, and the future MariaDB migration plan.

## 1. Deleted SharedView.html, unified share links into the main Detail page

**Problem**: Public share links (`?share=token`) rendered a fully separate ~1000-line standalone template (`SharedView.html`) duplicating most of Detail.html/Detail_js.html/Sheets_js.html. Editable share links could delete transactions/slips outright, which should never be allowed.

**Design**:
- `doGet` now always renders `Index` (the SPA shell). For `?share=token`, the validated token + its permission are embedded as `SHARE_TOKEN`/`SHARE_PERMISSION` inline JS globals in `Index.html`, read by `Shared_js.html` before boot.
- New `startShareApp()` in `Shared_js.html`: shows `#appView` directly, skips `#authView`/login entirely, hides bottom nav, loads the event via the same Detail rendering code.
- Server: extracted `_buildDetailPayload(ss, eventId, accountId)` — shared core used by both `getDetailData` (auth) and `getSharedEventView` (share), fixing a prior discrepancy where share visitors didn't see settlement "paid" state.
- Client permission model: single global `PERM = {canAdd, canEdit, canDelete, canDownload, canSummary, canShare}`. Owner: all true. Share view-only: only download/summary. Share edit: adds add/edit, but `canDelete` is **always false** regardless of link permission.
- Backend lockdown: **deleted `deleteDetailViaShare` and `deleteTransactionSlipViaShare` entirely** (not just hidden client-side) — an existing function is callable via devtools regardless of UI, so removing it server-side is the only real fix.
- `uploadTransactionSlipViaShare` stays (adding/replacing a photo is part of Add/Edit, not delete).
- Fixed an unrelated bug found in passing: `_saveTxSlipsIfNeeded` (Sheets_js.html) silently dropped failed upload/delete responses without a toast — now surfaces the error.

**Files touched**: `Code.js`, `Index.html`, `Shared_js.html`, `Detail_js.html`, `Sheets_js.html`, deleted `SharedView.html`.

## 2. Moved slip photo storage from Sheets cells to Google Drive (support ~500KB photos)

**Problem**: Google Sheets caps a cell at ~50,000 characters (~36KB after base64 overhead), far too small for a real receipt photo. User wanted up to ~500KB support.

**Design evolution** (three failed attempts before landing on the working one):
1. ~~Direct Drive embed URLs~~ (`lh3.googleusercontent.com/d/<id>`, `drive.google.com/thumbnail`, `/uc?export=`) — all **unreliable when hotlinked cross-origin** from inside Apps Script's sandboxed content iframe. Confirmed broken by live testing even though the URLs work fine when opened directly.
2. ~~`doGet` streaming the Drive file's raw `Blob`~~ — confirmed via live testing that Apps Script's `doGet` **only supports returning `HtmlOutput`/`TextOutput`**, not raw binary ("Script completed but the returned value is not a supported result type").
3. **Working solution**: store only a Drive `fileId`. Client fetches the image bytes via the same `google.script.run` bridge every other read in the app already uses (`getSlipImage(token, fileId)` / `getSlipImageViaShare(shareToken, fileId)`), then builds a `data:` URI locally. Cached client-side by fileId (`_slipSrcCache` in `Shared_js.html`).
- Server: `_uploadSlipToDrive(dataUri)` creates a **private** Drive file (no public sharing needed at all, since access goes through the authenticated RPC layer, gated by `_isKnownSlipFile(fileId)` checking the file actually appears in `TransactionSlips`).
- `_trashSlipFilesForTx`/`_trashSlipFilesForTxSet` added so deleting a transaction/event also cleans up its Drive file(s), not just the sheet row.
- Client compression (`Sheets_js.html`): single-pass `_compressSlipForUpload` targeting ~670,000 base64 chars (~500KB), replacing the old two-copy (preview + hi-res) approach that was capped much lower for the Sheets cell limit.
- `appsscript.json`: added explicit `oauthScopes` (`spreadsheets`, `drive`) — required the script owner to manually re-authorize once via the Apps Script editor (Run any function → grant Drive permission).

**Files touched**: `Code.js`, `Shared_js.html`, `Detail_js.html`, `Sheets_js.html`, `appsscript.json`.

## 3. Safari session-persistence investigation (long back-and-forth, ended without full resolution)

**Original complaint**: Safari forces re-login every time the app is closed and reopened, regardless of the configured `sessionMinutes` (Admin > Settings).

**Root cause (confirmed)**: Google Apps Script web apps are *always* served through a wrapper page that embeds the actual app content in a sandboxed iframe on a different origin (inherent to the platform, not fixable in app code). Safari's ITP (Intelligent Tracking Prevention) treats this as third-party/cross-site content and wipes `localStorage` (and equally, `sessionStorage`/`IndexedDB`/cookies) for it — confirmed to happen on every real app close, not just after Safari's usual 7-day cap. "Add to Home Screen" does **not** fix this, since the content still loads through the same sandboxed iframe regardless of launch method (confirmed by the user's own test).

**Attempts, in order, and what happened to each**:
1. **`?tk=<token>` URL + manual "Save Login on This Device" button** (Account menu → copy a login link → user manually bookmarks/Adds to Home Screen). This is the only mechanism confirmed *theoretically* sound (the URL itself carries the credential, sidestepping storage entirely) — but the user asked to remove this button/flow because they didn't want a manual "generate and save a link" step.
2. **Auto-redirect after login** (`top.location.href = <tk-url>` right after a successful "Remember me" login, no manual step) — implemented, but **confirmed by live user testing that the URL never changes at all**. Apps Script's sandboxed iframe blocks top-level navigation outright (not merely a user-activation-timing issue as initially hypothesized — it simply doesn't work, tested with no exceptions thrown and no navigation happening).
3. **Plain "Remember me" checkbox + localStorage vs sessionStorage** (no URL trick at all) — implemented per user's explicit request to simplify, but this **reintroduces the original bug** on Safari specifically (still gets wiped on close). Works fine on non-Safari browsers (Chrome/Android/Desktop) respecting the configured session length properly.

**Where this was left off**: the *only* proven-reliable mechanism remains the manual "get my login link, copy it, Add to Home Screen from that link" flow (attempt #1) — which the user had asked to remove. Last message from the assistant laid out the exact step-by-step user flow for bringing that button back; **the user had not yet confirmed whether to re-add it** before the conversation moved to the migration-planning topic. Current deployed state (v96) has the auto-redirect code in place (harmless — it silently no-ops) plus the Remember Me checkbox, but **no working fix for the core Safari complaint**.

**Minor bug fixed along the way**: after first adding the auto-redirect, login appeared to hang with no error if the top-navigation was blocked (early `return` skipped the fallback `startApp()` call) — fixed by making the redirect best-effort and always falling through to starting the app locally.

**Unresolved decision point for next session**: does the user want to re-add the manual "Save/Copy Login Link" button (only proven-working option), or accept the Safari-specific limitation and move on (especially now that a full backend migration is being considered, which would fix this permanently via real HTTP sessions/cookies)?

## 4. Future direction discussed: migrate off Apps Script + Sheets to HTML/JS/CSS + MariaDB

Motivated by wanting real performance (Sheets reads/writes the whole range every request) and to permanently fix the Safari session problem (real HTTP session cookies have no sandboxed-iframe/third-party-storage issue at all).

**Assessed scope** — roughly "rewrite the backend entirely, reuse most of the frontend":
1. **Stack decisions**: Node.js + Express recommended (same language as existing client code) or PHP/Python; hosting on a VPS or managed platform (not free like Apps Script); MariaDB self-hosted or managed; Drive-based slip storage → local disk or S3-compatible object storage (Cloudflare R2/Backblaze B2); need own domain + SSL (Let's Encrypt).
2. **Schema design**: existing "sheets" map directly to tables — `accounts`, `friends`, `events`, `details`, `event_friends`, `event_shares`, `sessions`, `transaction_slips`, `settlement_payments`. Add indexes on `eventId`/`accountId`/`token` — the main source of the speed win vs. Sheets' full-range reads.
3. **Data migration**: one-time script exporting Sheets data (via Sheets API or CSV) into MariaDB, plus migrating Drive-stored slip photos to the new storage.
4. **Backend rewrite**: port every function in `Code.js` (~1700 lines) to REST endpoints (`POST /api/login`, `GET /api/events`, `POST /api/details`, etc.) — this is the bulk of the migration effort.
5. **Auth redesign**: real HTTP session cookies (`httpOnly`, `secure`, `sameSite`) replacing the token-in-localStorage/URL workarounds entirely — this is what actually fixes the Safari problem for good, since there's no cross-origin iframe involved anymore.
6. **Client changes**: replace `gsc()`/`google.script.run` calls with `fetch()` against the new REST API — most UI/CSS code in `Detail_js.html`, `Sheets_js.html`, etc. is reusable as-is, only the RPC layer changes.
7. **Deploy/cutover**: basic CI/CD, run new stack in parallel with the Apps Script version during testing, have a rollback plan before flipping DNS/production traffic over.
8. **Ongoing maintenance** (things Apps Script currently handles for free): DB backups, security patching, uptime monitoring, SSL renewal.

**Status**: user asked for this plan; no implementation started. Next step floated: either design the DB schema concretely from the current sheet structures, or scaffold the new backend project — user has not yet chosen which to start with.

---

## Deploy history this session (ShareUp Apps Script project)
Versions 86 → 96, each `clasp push` + `clasp deploy -i <pinned-deployment-id>` + a matching git commit. Pinned deployment ID: `AKfycbzx0MuTHhC6g-bL6YhkB3kWoIxOSGBHmr-vDE7K1vBrwr2iie0ZL12MYCJautSGDOSi8w`. Local `clasp`/`node` note: the default `node` in PATH is v10 (incompatible with clasp) — had to invoke via the full v20 binary path explicitly: `/Users/admin/.nvm/versions/node/v20.17.0/bin/node /Users/admin/.nvm/versions/node/v20.17.0/bin/clasp ...`.
