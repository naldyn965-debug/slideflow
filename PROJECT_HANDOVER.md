# SlideFlow Phase 1: handover

## Status

Built and verified in headless Chromium (Local mode, strict CSP): 42 automated checks pass.
**Not verified**: live Firebase (auth, Firestore, Storage), the security rules in the emulator, the CSP against live Firebase endpoints, real-world PowerPoint/Keynote/Google Slides files. See `TESTING.md` for the manual list. Treat the Firebase path as written-to-spec but untested until that list is done.

## Architecture

No build. Browser ES modules, hash router, one global `ctx = { store, toast, user }` passed to views.

| Layer | Files | Notes |
|---|---|---|
| Storage contract | `js/store/index.js` | Views only talk to `ctx.store`. Add a backend by implementing the same shape. |
| Adapters | `store/local.js`, `store/firebase.js` | Chosen at boot by `useFirebase` in `config.js`. Firebase SDK loads from gstatic only in Firebase mode. |
| Parser | `pptx/parser.js` | Read-only. Returns `{slides[], width, height, warnings, zip}`. `summarize()` yields the JSON stored with the template. |
| Validation | `pptx/validate.js` | Extension, size, magic bytes (ZIP vs OLE), then full parse. Throws `PptxError` with a user-facing message and `code`. |
| UI | `ui/*`, `views/*` | `h()` builds DOM with text nodes only (no `innerHTML` for data). Native `<dialog>`. |

## Data model

Firestore `users/{uid}/templates/{id}`: `name, fileName, size, storagePath, slideCount, width, height (EMU), mediaCount, imageRefs, slideTitles[], warnings[], schemaVersion, createdAt, updatedAt`. Immutable after create.
Firestore `users/{uid}/projects/{id}`: `name, templateId|null, templateName|null, schemaVersion, createdAt, updatedAt`. `templateName` is a snapshot so a project survives its template being deleted.
Storage `users/{uid}/templates/{id}.pptx`. IDs are `crypto.randomUUID()`. Timestamps are client `Date.now()` numbers.

Only a small summary is stored in Firestore (well under the 1 MB document limit). Full parse output is recomputed from the file when previewing, which avoids storing large JSON and keeps the file the single source of truth.

## Security

- Per-user isolation and field validation in `firebase/*.rules`; default deny elsewhere.
- Firebase web config is public; it is not a secret. No server secrets exist in this repo. Keep it that way: Phase 2 AI keys must live in a server function, never in `js/`.
- Uploaded files are only parsed as data. XML parts are size-capped, total uncompressed size is capped (zip-bomb guard), and nothing from the file is executed or inserted as HTML. SVG media is shown via `<img>` only.
- CSP in `vercel.json` is strict (`style-src 'self'`, no inline script). Written from the Firebase v10 SDK's known endpoints but not run against a live project; loosen `connect-src`/`frame-src` if the console shows violations.
- Chromium logs a harmless CSP console message when it renders its own error page for malformed XML; it only happens for invalid uploads.
- Guest (anonymous) accounts are tied to the browser. Upgrade path: `linkWithPopup` to attach Google to the guest uid.

## Cost and performance

Static hosting, no functions. Costs come from Storage and Firestore reads: list queries cap at 200 docs; previews download the file once per open. Parsing is client-side. Limits (50 MB, 200 slides) are in `config.js` and mirrored in the rules; change both together.

## Known limitations

- Preview is structural, not a faithful render (see README).
- Group and placeholder geometry ignore rotation and flips. Placeholders whose position cannot be resolved are listed but not drawn.
- Local-mode data is not migrated to Firebase.
- No rename or edit for projects or templates; no multi-file upload; no search or pagination beyond 200 items.
- Hash routing; Firebase Hosting or Vercel rewrites are not needed.

## Phase 2 entry points

- **Slide model / editor**: start from `parsePptx().slides` (shapes with boxes, text, image refs). Add a normalized `Slide` schema in a new `js/model/` and store it on the project (`projects/{id}`) or a subcollection; bump `schemaVersion`.
- **AI generation**: add a server endpoint (Vercel function) that holds the key; the client calls it with project and template summary. Do not call model APIs from the browser.
- **Export**: JSZip is already vendored; writing back to PPTX should patch the original template's XML parts rather than regenerate the file, to preserve masters and themes.
- **Project page** (`views/project.js`) is the mount point for the editor.
