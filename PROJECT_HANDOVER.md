# SlideFlow handover (Phase 1 + Phase 2 + Phase 2.1 + Phase 2.2 + Phase 2.5 Step 3)

## Status

Phases 1 and 2 were verified in headless Chromium (Local mode, strict CSP): 92 e2e checks (42 Phase 1, 50 Phase 2). Phase 2.1 (precision fixes, below) was built in a sandbox without Chromium: its 55 browser-free analysis checks pass (`npm run test:analysis`), its e2e checks were not run then. Phase 2.2 (below) was verified end to end: `npm test` 98/98 in headless Chromium and 96 analysis checks (the analysis suite was run through Chromium's DOMParser because jsdom could not be installed in that sandbox; run `npm run test:analysis` once with jsdom).
Phase 2.5 Step 3 (size-adaptive floors, design version 4, below): `npm test` 108/108 in headless Chromium; analysis 206/206 without the real FOOD deck and 237/237 with it, both run through Chromium's DOMParser because jsdom was unavailable in the authoring sandbox; a one-off version 3 vs version 4 regression comparison reported 28 of 31 decks unchanged except `version` and 3 changed as expected; 7 scaled-threshold mutations were all detected. **Still pending for Step 3**: a native jsdom run of `npm run test:analysis` and the Firebase manual checklist (both listed in `TESTING.md`).
**Not verified**: live Firebase (auth, Firestore, Storage), the security rules in the emulator, the CSP against live Firebase endpoints, real-world PowerPoint/Keynote/Google Slides files. See `TESTING.md` for the manual list. Treat the Firebase path as written-to-spec but untested until that list is done.

## Architecture

No build. Browser ES modules, hash router, one global `ctx = { store, toast, user }` passed to views.

| Layer | Files | Notes |
|---|---|---|
| Storage contract | `js/store/index.js` | Views only talk to `ctx.store`. Add a backend by implementing the same shape. |
| Adapters | `store/local.js`, `store/firebase.js` | Chosen at boot by `useFirebase` in `config.js`. Firebase SDK loads from gstatic only in Firebase mode. |
| Parser | `pptx/parser.js` | Read-only. Returns `{slides[], width, height, warnings, zip}`. A shape's own rotation is exposed as `rot` (degrees, absent when 0); the preview turns the unrotated box by it. A parent group's rotation is not read. `summarize()` yields the JSON stored with the template. |
| Design analysis | `pptx/design.js` | `extractDesign(parsed)` returns the Design System JSON. Reuses parser helpers (exported) and re-reads slide XML; Phase 1 parsing is unchanged. |
| Validation | `pptx/validate.js` | Extension, size, magic bytes (ZIP vs OLE), then full parse. Throws `PptxError` with a user-facing message and `code`. |
| UI | `ui/*`, `views/*` | `h()` builds DOM with text nodes only (no `innerHTML` for data). Native `<dialog>`. |

## Phase 2: Design System

Template record gains `design` (object), `designVersion` (int; `DESIGN_VERSION` was 2 when Phase 2 shipped and is now 4, see Phase 2.5 Step 3) and `mode` (`preserve` | `enhance`, absent = preserve). Projects gain `mode`. All other fields, IDs and file paths are untouched. Analysis updates never change `updatedAt`, so library order is stable.

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
- **Classification** (`classify()`): order is chart/table, title (first slide: a title layout, or little text with at most 3 further text boxes, or any number of them when the title is at least 1.5x larger than all other text), closing (wording or layout name in last two slides, or a last slide framed like the cover), section (layout type or very short title-only slide), comparison (two side-by-side text blocks), image (content pictures cover 30% or more, or a full-slide picture with almost no text), text, other. Layout type from the file (`secHead`, `twoTxTwoObj`...) gives `high`; content-only evidence gives `medium`. A picture covering the whole slide is a *background* (`layoutRole: background`) and never counts as image content.
- **Smart Enhance** (`buildModes()`): body at least 16 pt, title at least 28 pt and 1.25x body, footers at least 12 pt, margins at least 5/5/5/4 %, text colour swapped for the theme colour with better contrast when under 4.5:1. Fonts and palette are never replaced. Since design version 4 the three point minimums (16 / 28 / 12 pt) are the values for a 13.333 x 7.5 in slide and are multiplied by the slide-size scale `s` (see Phase 2.5 Step 3); margins and the 4.5:1 ratio are not scaled.

## Phase 2.1: precision fixes (found by testing a real dark template)

Symptoms and root causes, all in `design.js`; no architecture change, `design` stays additive (all Phase 2 fields kept, new ones optional). `DESIGN_VERSION` is now 2, so stored v1 analyses are re-run automatically on the next preview (cards keep working with v1 data until then).

| Symptom | Cause | Fix |
|---|---|---|
| Low contrast on every slide although text is white on #211D19 | Colour-map overrides were ignored. Dark decks written by PowerPoint put `overrideClrMapping` (bg1=dk1, tx1=lt1) on the slide or layout, so unset text colour resolved to the *theme's dark* colour. Also: a text box was only compared with the slide background, ignoring panels, translucent fills and pictures under it; autoshape text colour from `<p:style><a:fontRef>` was ignored; fill and text alpha ignored | Effective colour map = master, then layout override, then slide override (`clrMapOf`). Contrast now uses `backdropOf()`: slide background, then master/layout shapes, earlier slide shapes and the shape's own fill, composited in z-order with alpha. Pictures, gradients and effect fills give no single colour, so that text is reported as *not evaluated* instead of guessed. Partial overlaps and non-rectangular shapes add candidate colours; text is flagged only when it fails against every candidate. `fontRef` colour added to the style chain |
| Negative margins | Content bounds used raw boxes, so bleeding pictures and off-slide decoration gave negative values | `roleOf()` labels each shape `content`, `bleed`, `background`, `footer` or `offslide`. Margins use content only, clipped to the slide and never negative. Text boxes that truly run off the slide set `offSlide` |
| Slides 1 and 15 classed as image | Image coverage was tested before position, and a full-slide picture counted as image content | Pictures covering the slide are backgrounds. First slide is a title by position and wording. Last slide with a full-slide picture and little text is a closing (high when it mirrors the cover). Layout-name hints (`Title Slide`, `Thank you`, `Contact`). A lone large text box is inferred as the title so dividers are found |

Diagnostics: each slide record carries `colorMap` (`master`/`layout`/`slide`), `contrast` (`ok`, `low`, `not-evaluated`, `no-text`) with the worst pair, `bleed`, `offSlide` and `background.picture`; the `low-contrast` finding lists up to 5 `evidence` entries (slide, text colour, backdrop colour, ratio, first 30 characters). The Design tab and the slide details show them.

Known limits kept honest: the real template was not available, so the fixes were driven by a synthetic deck that reproduces each symptom (`tests/make_fixture_p21.py`); if the real file still misbehaves, its design JSON now says why. Bleed detection treats any non-text shape crossing an edge, and any picture touching an edge, as decoration. Text over a picture is never judged for contrast. Slide-level `masterClrMapping` is treated as "inherit from the layout".

Tuning knobs worth knowing: thresholds in `classify()` and `buildModes()`; the point floors in `sizeFloors()` (Phase 2.5 Step 3); `DESIGN_VERSION` (bump to force re-analysis of all stored templates on next preview).

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

## Phase 2.5 Step 3: size-adaptive floors (design version 4)

Why: every point floor used to be a constant written for a 13.333 x 7.5 in slide, so a 20 x 11.25 in deck (the real FOOD template) was judged by thresholds that are tiny for its page, and a 10 x 5.625 in deck by thresholds that are large for its page. All changes are in `js/pptx/design.js` (`sizeScale()`, `sizeFloors()`, `DESIGN_VERSION`); the design JSON shape is unchanged (additive rules only), and no store, rules or UI code changed except that the stored-version check in `js/ui/preview.js` now sees 4.

**Approved decisions**
1. `wr = width / 12192000` and `hr = height / 6858000`.
2. The formula below and `DESIGN_VERSION = 4`.
3. On a 10 x 5.625 in slide a 10 pt footer stays unchanged, because the scaled floor is 9 pt.
4. FOOD's body text change from 21 pt to 24 pt is intentional.

**Formula**

```
s = clamp( min( max(w / 12192000, h / 6858000), 1.2 * sqrt(wr * hr) ), 0.25, 4 )
```

`max(wr, hr)` follows the larger side. The `1.2 * sqrt(wr * hr)` term keeps an extreme aspect ratio (32:9 banner, portrait) from scaling by its long side alone. Width or height that is zero or not a number gives `s = 1` (the parser rejects such sizes anyway). Each floor is `reference * s`, rounded to 0.1 pt (the precision run sizes are kept at). 13.333 x 7.5 in and 10 x 7.5 in (4:3) give `s = 1`: the reference floors.

**What the scale factor does** (reference floors at `s = 1`)

| Threshold | Reference | Where it is used | 10 x 5.625 in (s 0.75) | 20 x 11.25 in (s 1.5) | 26.67 x 7.5 in (s 1.697) |
|---|---|---|---|---|---|
| Body | 16 pt | Smart Enhance body minimum | 12 | 24 | 27.2 |
| Title | 28 pt | Smart Enhance title minimum, together with `round(1.25 x scaled body)` | 21 | 42 | 47.5 |
| Subtitle | none | Smart Enhance keeps it at least as large as the scaled body | follows body | follows body | follows body |
| Footer | 12 pt | Smart Enhance footer minimum | 9 | 18 | 20.4 |
| Small text | 12 pt | `small-text` finding: a run below this adds the slide | 9 | 18 | 20.4 |
| Low body size | 14 pt | `low-body-size` finding (info): dominant body below this | 10.5 | 21 | 23.8 |
| Large text | 18 pt | contrast class: a regular run at or above this needs 3:1, below it 4.5:1 | 13.5 | 27 | 30.5 |
| Large bold text | 14 pt | contrast class: a bold run at or above this needs 3:1 | 10.5 | 21 | 23.8 |

Not scaled: the 4.5:1 and 3:1 contrast ratios, the margins (a percentage of the slide; Smart Enhance floors 5 / 5 / 5 / 4 %), the 1.25 title-to-body ratio itself (its input, the body size, is scaled), the 4 % margin-consistency tolerance and the 110-word limit. Rules are minimums: a value at or above its floor never changes, Smart Enhance never lowers anything, and the Preserve -> Enhance differences still equal `Enhance.changes` exactly. When `s` is not 1 the change reasons and the two finding messages append "The N pt reference minimum is scaled xS for this slide size."; at `s = 1` no such note is added.

**Behaviour changes on existing decks**: the pptxgenjs fixture (10 x 5.625 in) no longer gets a footer change (10 pt >= 9 pt); `p25-food-like.pptx` and the real FOOD deck (20 x 11.25 in) now get body 21 -> 24 pt (and nothing else). Decks at 13.333 x 7.5 in or 10 x 7.5 in analyse as before.

**Migration and reanalysis (`DESIGN_VERSION = 4`)**
- `js/ui/preview.js` uses the stored design only when `t.designVersion === DESIGN_VERSION`. A template saved at version 1, 2 or 3 is analysed again from its stored file the next time its Preview opens (or when Re-analyze is pressed), and `ctx.store.templates.update(t, { design, designVersion })` saves the result.
- That patch contains only `design` and `designVersion`: ID, name, file, `mode`, `createdAt` and `updatedAt` are untouched (library order is stable). The Firestore update rules already allow exactly these fields, so no rules change or redeploy is needed beyond what Phase 2 required.
- Until a template's Preview is opened, library cards show its stored (older) analysis. If saving fails, the Preview shows the fresh analysis, a toast reports the failure, and the saved record keeps its old version, so the next Preview tries again.
- Design JSON downloaded before the upgrade says `version: 3` and is not rewritten. Phase 3 should read `design.version >= 4` before trusting the point tokens for non-standard slide sizes.

**Tests** (details and results in `TESTING.md`)
- `tests/make_fixture_p25s3.py`: 13 `p25-s3-*.pptx` decks (sizes 6.67 x 3.75, 13.33 x 7.5, 20 x 11.25, 26.67 x 7.5 and 7.5 x 10.83 in), with contrast, margin and exactly-on-the-floor variants. The expected scales and floors in `tests/analysis.mjs` are hard-coded literals from an independent calculation.
- Reported results: e2e 108/108 (headless Chromium); analysis 206/206 without FOOD, 237/237 with the real FOOD deck (31 FOOD-only checks); regression comparison 28 of 31 decks unchanged except `version`, 3 changed as expected (the three changed decks are the pptxgenjs fixture, `p25-food-like.pptx` and FOOD; the comparison script and the list of 31 decks are not in the repository, so the total could not be reproduced from the repository alone); mutation checks: 7 one-off mutations, one per scaled threshold, all detected (not a committed script).
- Execution limits: jsdom was unavailable in the authoring sandbox, so the analysis suite ran through Chromium's DOMParser (adapter used as the alternative). **Native jsdom tests are still pending**, and the **Firebase manual checks are still pending** (the checklist was not run). Neither is claimed as passed, including the version 3 -> 4 save against Firestore (the e2e migration check covers Local mode only).

**Known limitations**
- `inferTitle()` still uses a fixed 28 pt rule: a slide with no title placeholder and a single text candidate gets that box as its title only if its largest run is 28 pt or more. It is not scaled with the slide size, so on slides much smaller or larger than 13.333 x 7.5 in it can be expected to fire too late or too early (not measured).
- The `inconsistent-title-size` finding still uses a fixed 4 pt tolerance against the median title size of text slides. It is not scaled either, so on a large slide a 4 pt spread is a smaller relative difference than on a small one.
- Both are candidates for a later step; neither was changed here on purpose (approved scope).

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
- Analysis of legacy templates happens on first Preview and writes the result back (derived cache; no existing field is changed). A version bump (v1 to v2, and most recently v3 to v4) triggers the same refresh.
- The Firestore rules for Phase 2 (update path) are untested against the emulator, as is the version 3 to 4 re-analysis write against live Firestore.
- `inferTitle` (28 pt) and `inconsistent-title-size` (4 pt) are fixed values and do not follow the slide size (Phase 2.5 Step 3).
- The analysis suite has been run through Chromium's DOMParser only; a native jsdom run of `npm run test:analysis` and the Firebase manual checklist are pending.
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
