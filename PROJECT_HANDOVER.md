# SlideFlow handover (Phase 1 + Phase 2 + Phase 2.1 + Phase 2.2)

## Status

Phases 1 and 2 were verified in headless Chromium (Local mode, strict CSP): 92 e2e checks (42 Phase 1, 50 Phase 2). Phase 2.1 (precision fixes, below) was built in a sandbox without Chromium: its 55 browser-free analysis checks pass (`npm run test:analysis`), its e2e checks were not run then. Phase 2.2 (below) was verified end to end: `npm test` 98/98 in headless Chromium and 96 analysis checks (the analysis suite was run through Chromium's DOMParser because jsdom could not be installed in that sandbox; run `npm run test:analysis` once with jsdom).
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

Template record gains `design` (object), `designVersion` (int, currently `DESIGN_VERSION = 2`) and `mode` (`preserve` | `enhance`, absent = preserve). Projects gain `mode`. All other fields, IDs and file paths are untouched. Analysis updates never change `updatedAt`, so library order is stable.

Shape of `design` (`schema: "slideflow.design"`):

```
slide {width,height,aspect}            theme {name, colors{dk1..accent6}, fonts{major,minor}}
palette {colors[{hex,slot,neutral,roles,weight}], offThemeCount}
typography {fonts[{name,share}], sizes[{pt,share}], roles{title,subtitle,body,footer:{font,sizePt,bold,color}}, minPt,maxPt,boldShare}
backgrounds {types, colors[]}          shapes {total, byKind, presets}
layouts {masters[{part,theme,layouts,decor[]}], list[{name,type,used,decor,placeholders[{type,box%}]}]}
spacing {marginsPct{left,right,top,bottom}, spreadPct, gapPct, textInsetsEmu, alignment{left..}}
patterns[{id,layout,slides[],class}]   recurring[{kind,label,box,count}]
classification {counts, confidence}    slides[{n,title,layout,layoutType,class,confidence,reasons[],background{type,hex,source,picture?},words,counts,imageCoverage,contentImageCoverage,bounds,bleed[]?,offSlide?,colorMap,contrast{status,ratio?,text?,background?,unchecked?},minPt,maxPt}]
modes {preserve{tokens}, enhance{tokens,changes[],findings[]}, note}     limitations[]
```
Boxes are percentages of the slide. No nested arrays (Firestore restriction); the test suite asserts this. Typical size 6-12 KB.

How it decides (all in `design.js`, each rule commented):
- **Style resolution**: run properties, then shape `lstStyle`, layout placeholder, master placeholder, master `txStyles` (title/body/other), then `presentation.xml` defaults. `+mj-lt`/`+mn-lt` resolve via the theme. RTL runs use the `cs` font.
- **Colours**: `srgbClr`, `schemeClr` (through the master colour map), `sysClr`; lumMod/lumOff applied, other transforms flagged approximate.
- **Classification** (`classify()`): order is chart/table, title (first slide), closing (wording or layout name in last two slides, or a last slide framed like the cover), section (layout type or very short title-only slide), comparison (two side-by-side text blocks), image (content pictures cover 30% or more, or a full-slide picture with almost no text), text, other. Layout type from the file (`secHead`, `twoTxTwoObj`...) gives `high`; content-only evidence gives `medium`. A picture covering the whole slide is a *background* (`layoutRole: background`) and never counts as image content.
- **Smart Enhance** (`buildModes()`): body at least 16 pt, title at least 28 pt and 1.25x body, footers at least 12 pt, margins at least 5/5/5/4 %, text colour swapped for the theme colour with better contrast when under 4.5:1. Fonts and palette are never replaced.

## Phase 2.1: precision fixes (found by testing a real dark template)

Symptoms and root causes, all in `design.js`; no architecture change, `design` stays additive (all Phase 2 fields kept, new ones optional). `DESIGN_VERSION` is now 2, so stored v1 analyses are re-run automatically on the next preview (cards keep working with v1 data until then).

| Symptom | Cause | Fix |
|---|---|---|
| Low contrast on every slide although text is white on #211D19 | Colour-map overrides were ignored. Dark decks written by PowerPoint put `overrideClrMapping` (bg1=dk1, tx1=lt1) on the slide or layout, so unset text colour resolved to the *theme's dark* colour. Also: a text box was only compared with the slide background, ignoring panels, translucent fills and pictures under it; autoshape text colour from `<p:style><a:fontRef>` was ignored; fill and text alpha ignored | Effective colour map = master, then layout override, then slide override (`clrMapOf`). Contrast now uses `backdropOf()`: slide background, then master/layout shapes, earlier slide shapes and the shape's own fill, composited in z-order with alpha. Pictures, gradients and effect fills give no single colour, so that text is reported as *not evaluated* instead of guessed. Partial overlaps and non-rectangular shapes add candidate colours; text is flagged only when it fails against every candidate. `fontRef` colour added to the style chain |
| Negative margins | Content bounds used raw boxes, so bleeding pictures and off-slide decoration gave negative values | `roleOf()` labels each shape `content`, `bleed`, `background`, `footer` or `offslide`. Margins use content only, clipped to the slide and never negative. Text boxes that truly run off the slide set `offSlide` |
| Slides 1 and 15 classed as image | Image coverage was tested before position, and a full-slide picture counted as image content | Pictures covering the slide are backgrounds. First slide is a title by position and wording. Last slide with a full-slide picture and little text is a closing (high when it mirrors the cover). Layout-name hints (`Title Slide`, `Thank you`, `Contact`). A lone large text box is inferred as the title so dividers are found |

Diagnostics: each slide record carries `colorMap` (`master`/`layout`/`slide`), `contrast` (`ok`, `low`, `not-evaluated`, `no-text`) with the worst pair, `bleed`, `offSlide` and `background.picture`; the `low-contrast` finding lists up to 5 `evidence` entries (slide, text colour, backdrop colour, ratio, first 30 characters). The Design tab and the slide details show them.

Known limits kept honest: the real template was not available, so the fixes were driven by a synthetic deck that reproduces each symptom (`tests/make_fixture_p21.py`); if the real file still misbehaves, its design JSON now says why. Bleed detection treats any non-text shape crossing an edge, and any picture touching an edge, as decoration. Text over a picture is never judged for contrast. Slide-level `masterClrMapping` is treated as "inherit from the layout".

Tuning knobs worth knowing: thresholds in `classify()` and `buildModes()`; `DESIGN_VERSION` (bump to force re-analysis of all stored templates on next preview).

## Phase 2.2: slide types and geometry (design version 3)

All in `design.js` (plus two UI wording lines in `design-panel.js` / `preview.js`). `DESIGN_VERSION` is 3, so stored analyses refresh on the next preview. Every earlier field is kept; new fields are optional.

**Classification**
| Before | Now |
|---|---|
| Any two side-by-side text blocks = comparison | Two columns are a layout choice. Comparison needs its layout (`twoTxTwoObj` / name "compar...", high), wording in the title (high) or in a short label such as "vs" (medium), or mirrored columns that each have their own short heading (`mirrored()`, medium). Running prose ("compared with last year") is not wording. Plain two-column text stays `text`, and its reasons say so |
| Picture coverage >= 30% = image, or >= 15% with <= 60 words | Only *content* pictures count (`pictureRole()`): backgrounds, logos/icons (< 3% of the slide), thin bands/rails (< 12% thick) and pictures repeated at the same place on >= 30% of slides (template artwork) are decorative. Image needs light supporting text: coverage >= 30% with <= 30 body words, a picture layout with <= 40, a gallery of 3+ with <= 40, or >= 15% with a caption <= 15. A picture beside a paragraph stays `text` (medium confidence when the picture competes) |

New slide fields: `imageRoles {background, content, decorative}`; `contentImageCoverage` now counts content pictures only (`imageCoverage` still counts all).

**Geometry** (root causes of "everything is bleed")
| Cause | Fix |
|---|---|
| Shape rotation was never read: a vertical rule or side label built from a rotated horizontal shape was measured unrotated, so it appeared to cross the edge | `rotOf()` / `rotatedBox()` apply a shape's own rotation to its bounds (group rotation is still ignored) |
| Tables, charts and diagrams past the edge were labelled bleed, hidden from margins and never flagged | They stay `content`; past the edge they are overflow |
| A text frame overshooting by less than its own inset was flagged off-slide | Text is judged by frame minus insets |
| Empty placeholders were treated as decoration | New role `empty` (not drawn when presenting): neither bleed nor overflow |
| No evidence of intent, so a bar or band repeated on every template slide read as a problem | `bleed` = decorative, intentional (reported, not a problem). `bleedDetail` lists `{kind, edges, mode flush|beyond, byPct, repeats}`; `repeats` = number of slides carrying the same shape at the same place |

Overflow is `offSlide` + `overflow [{kind text|table|chart|object, edges, byPct, outside?, sample?}]`, summarised by `spacing.offSlideSlides` and a new `content-overflow` finding. Decorative bleed is `bleed` / `bleedDetail` / `spacing.bleedSlides`. Not yet seen: the real template. If every slide still reports bleed, check `bleedDetail`: a `repeats` equal to the slide count means a template element, which is correct.
Tests: `tests/make_fixture_p22.py` (4 fixtures), 41 new analysis checks, 2 e2e checks.

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
- Analysis of legacy templates happens on first Preview and writes the result back (derived cache; no existing field is changed). A version bump (v1 to v2) triggers the same refresh.
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
