# SlideFlow handover (Phase 1 + Phase 2)

## Status

Phases 1 and 2 built and verified in headless Chromium (Local mode, strict CSP): 92 automated checks pass (42 Phase 1 regression, 50 Phase 2).
**Not verified**: live Firebase (auth, Firestore, Storage), the security rules in the emulator, the CSP against live Firebase endpoints, real-world PowerPoint/Keynote/Google Slides files. See `TESTING.md` for the manual list. Treat the Firebase path as written-to-spec but untested until that list is done.

## Architecture

No build. Browser ES modules, hash router, one global `ctx = { store, toast, user }` passed to views.

| Layer | Files | Notes |
|---|---|---|
| Storage contract | `js/store/index.js` | Views only talk to `ctx.store`. Add a backend by implementing the same shape. |
| Adapters | `store/local.js`, `store/firebase.js` | Chosen at boot by `useFirebase` in `config.js`. Firebase SDK loads from gstatic only in Firebase mode. |
| Parser | `pptx/parser.js` | Read-only. Returns `{slides[], width, height, warnings, zip}`. `summarize()` yields the JSON stored with the template. |
| Design analysis | `pptx/design.js` | `extractDesign(parsed)` returns the Design System JSON. Reuses parser helpers (exported) and re-reads slide XML; Phase 1 parsing is unchanged. |
| Validation | `pptx/validate.js` | Extension, size, magic bytes (ZIP vs OLE), then full parse. Throws `PptxError` with a user-facing message and `code`. |
| UI | `ui/*`, `views/*` | `h()` builds DOM with text nodes only (no `innerHTML` for data). Native `<dialog>`. |

## Phase 2: Design System

Template record gains `design` (object), `designVersion` (int, currently `DESIGN_VERSION = 1`) and `mode` (`preserve` | `enhance`, absent = preserve). Projects gain `mode`. All other fields, IDs and file paths are untouched. Analysis updates never change `updatedAt`, so library order is stable.

Shape of `design` (`schema: "slideflow.design"`):

```
slide {width,height,aspect}            theme {name, colors{dk1..accent6}, fonts{major,minor}}
palette {colors[{hex,slot,neutral,roles,weight}], offThemeCount}
typography {fonts[{name,share}], sizes[{pt,share}], roles{title,subtitle,body,footer:{font,sizePt,bold,color}}, minPt,maxPt,boldShare}
backgrounds {types, colors[]}          shapes {total, byKind, presets}
layouts {masters[{part,theme,layouts,decor[]}], list[{name,type,used,decor,placeholders[{type,box%}]}]}
spacing {marginsPct{left,right,top,bottom}, spreadPct, gapPct, textInsetsEmu, alignment{left..}}
patterns[{id,layout,slides[],class}]   recurring[{kind,label,box,count}]
classification {counts, confidence}    slides[{n,title,layout,layoutType,class,confidence,reasons[],background,words,counts,bounds,minPt,maxPt}]
modes {preserve{tokens}, enhance{tokens,changes[],findings[]}, note}     limitations[]
```
Boxes are percentages of the slide. No nested arrays (Firestore restriction); the test suite asserts this. Typical size 6-12 KB.

How it decides (all in `design.js`, each rule commented):
- **Style resolution**: run properties, then shape `lstStyle`, layout placeholder, master placeholder, master `txStyles` (title/body/other), then `presentation.xml` defaults. `+mj-lt`/`+mn-lt` resolve via the theme. RTL runs use the `cs` font.
- **Colours**: `srgbClr`, `schemeClr` (through the master colour map), `sysClr`; lumMod/lumOff applied, other transforms flagged approximate.
- **Classification** (`classify()`): order is chart/table, title (first slide), closing (wording in last two slides), section (layout type or very short title-only slide), comparison (two side-by-side text blocks), image (coverage of 30% or more), text, other. Layout type from the file (`secHead`, `twoTxTwoObj`...) gives `high`; content-only evidence gives `medium`.
- **Smart Enhance** (`buildModes()`): body at least 16 pt, title at least 28 pt and 1.25x body, footers at least 12 pt, margins at least 5/5/5/4 %, text colour swapped for the theme colour with better contrast when under 4.5:1. Fonts and palette are never replaced.

Tuning knobs worth knowing: thresholds in `classify()` and `buildModes()`; `DESIGN_VERSION` (bump to force re-analysis of all stored templates on next preview).

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
- Design analysis runs on the main thread (about 60 ms for 11 slides; yields every 10 slides). A 200-slide deck may take a couple of seconds; move to a Web Worker if that becomes noticeable.
- Analysis of legacy templates happens on first Preview and writes the result back (derived cache; no existing field is changed).
- The Firestore rules for Phase 2 (update path) are untested against the emulator.
- No rename or edit for projects or templates; no multi-file upload; no search or pagination beyond 200 items.
- Hash routing; Firebase Hosting or Vercel rewrites are not needed.

## Phase 3 instructions (AI slide generation)

Do not start by calling a model. Build in this order:

1. **Server boundary first.** Add a Vercel function (`/api/generate`) that verifies the Firebase ID token, holds the model key, rate-limits per user and caps tokens. The browser never sees the key.
2. **Input contract.** Send the outline/content plus, from the template: `design.modes[project.mode].tokens`, `design.classification`, and `design.slides` reduced to `{n, class, layout, words, counts}`. Never send the PPTX.
3. **Plan, then build.** Have the model return a plan only: an ordered list of `{class, title, bullets|columns|image prompt|chart data}`. Validate it in code (allowed classes, word limits per class from the template's own slides, at most N slides).
4. **Map plan to template.** For each planned slide pick an exemplar slide of that class (prefer `confidence: high`, and slides inside a `design.patterns` group), then fill its placeholders. Reuse the exemplar's layout name so masters, theme and fixed logos carry over. If a class has no exemplar, fall back to `text` and tell the user.
5. **Apply the mode.** Preserve: copy exemplar values unchanged. Enhance: apply `modes.enhance.tokens` (sizes, margins, text colour) when writing, and list what was changed.
6. **Write the file by patching XML** (JSZip is already vendored): clone the exemplar slide part and its rels, replace text runs, update `presentation.xml` `sldIdLst` and rels, and fix `[Content_Types].xml`. Do not regenerate the file from scratch. Keep this in `js/pptx/writer.js`, separate from `parser.js`.
7. **Verify by round trip.** Re-parse the output with `parsePptx` and `extractDesign`; compare slide count, classes and fonts with the plan. Fail loudly rather than ship a mismatch.
8. **Store results** as a new `projects/{id}/slides` subcollection or a project field with its own `schemaVersion`; extend `firestore.rules` in the same change.
9. **Tests**: add fixtures from real PowerPoint, Keynote and Google Slides exports first. Those are the biggest unknown and will probably adjust the classification thresholds.

Known follow-ups that make Phase 3 better: localized layout-name hints, picture pixel size, per-class exemplar scoring, and a Web Worker for analysis.

## Entry points for editor and export (later phases)

- **Slide model / editor**: start from `parsePptx().slides` (shapes with boxes, text, image refs). Add a normalized `Slide` schema in a new `js/model/` and store it on the project (`projects/{id}`) or a subcollection; bump `schemaVersion`.
- **AI generation**: add a server endpoint (Vercel function) that holds the key; the client calls it with project and template summary. Do not call model APIs from the browser.
- **Export**: JSZip is already vendored; writing back to PPTX should patch the original template's XML parts rather than regenerate the file, to preserve masters and themes.
- **Project page** (`views/project.js`) is the mount point for the editor.
