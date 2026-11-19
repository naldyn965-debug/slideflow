# SlideFlow, Phase 1 and 2

Dashboard, PPTX template management, a PPTX parser and (Phase 2) template intelligence. Plain HTML, CSS and ES modules. No build step, no framework.
Out of scope so far: AI generation, research, slide editor, export.

## What works

- **Dashboard**: sidebar, recent projects, template library, create project, loading / empty / error states.
- **Templates**: upload (click or drag and drop), validate, review, save, list, structural preview, delete.
- **Projects**: create (optionally from a template), list, open, delete. A project is a record pointing at a template; the editor comes later.
- **Parser** (`js/pptx/parser.js`): slide count and order, slide size, per-slide text, shape positions, layout names, hidden slides, image references (with the real picture shown in the preview when the browser supports the format).
- **Two storage modes**, same code path for the UI:
  - **Local mode** (default, no setup): IndexedDB in this browser.
  - **Firebase mode** (Auth + Firestore + Storage): switches on when `js/config.js` has a Firebase config.

## Phase 2: template intelligence

Every template is analysed in the browser into a versioned JSON **Design System** (`js/pptx/design.js`), stored on the template record and shown in the library and the preview's **Design** tab.

- **Extracted**: theme colours and fonts; colours actually used (text, fills, backgrounds) with theme slots; fonts, sizes, bold/regular and colour per role (title, subtitle, body, footer), with sizes resolved through the layout, master and presentation defaults; backgrounds per slide; shape and picture counts; margins, gaps and alignment from real positions; masters, layouts (with usage and fixed logos/bars); repeated slide structures and recurring logos or footers.
- **Slide classes**: title, section divider, text, image, comparison, chart or table, closing, other. Rule-based, each with a confidence (high / medium / low) and the reasons shown per slide.
- **Modes** (per template, chosen in the Design tab; projects copy it):
  - **Preserve Design**: tokens are exactly what was observed.
  - **Smart Enhance**: same fonts, palette and layouts, but sizes, margins and contrast are lifted to readable minimums. The point minimums are written for a 13.333 x 7.5 in slide and scale with the slide size (see "Size-adaptive thresholds" below). Findings (small text, low contrast, too many fonts, inconsistent margins...) and the exact list of changes are shown.
  Modes are a stored contract for Phase 3 slide generation. **Nothing modifies your PPTX file.**
- **Existing templates** (uploaded in Phase 1) keep their IDs and fields. They show "Design not analyzed yet"; opening Preview analyses them once and caches the result. Outdated analyses are replaced the same way: a design saved by an earlier engine version (for example version 3, before size-adaptive thresholds) is re-run with the current one the next time its Preview opens (see "Design version 4" below).
- **Download design JSON** from the Design tab.

### Size-adaptive thresholds (Phase 2.5 Step 3, design version 4)

Point sizes mean different things on different slide sizes: 16 pt is body text on a 13.333 x 7.5 in slide, but it is large text on a 10 x 5.625 in slide and small text on a 20 x 11.25 in slide. Since design version 4 every **point** threshold is multiplied by one scale factor `s` computed from the slide size (`sizeScale()` / `sizeFloors()` in `js/pptx/design.js`):

```
wr = slide width  / 12192000        (EMU; 13.333 in)
hr = slide height / 6858000         (EMU; 7.5 in)
s  = clamp( min( max(wr, hr), 1.2 * sqrt(wr * hr) ), 0.25, 4 )
```

`max(wr, hr)` follows the larger dimension. The `1.2 * sqrt(wr * hr)` term (the geometric mean of the two ratios) stops an extreme aspect ratio, such as a 32:9 banner or a portrait slide, from scaling by its long side alone. The result is limited to 0.25 to 4. A 13.333 x 7.5 in slide and a 10 x 7.5 in (4:3) slide both give `s = 1`, so their thresholds are the reference values. Each scaled value is rounded to 0.1 pt.

| Threshold (reference at s = 1) | What it controls | 10 x 5.625 in (s = 0.75) | 20 x 11.25 in (s = 1.5) |
|---|---|---|---|
| Body, 16 pt | Smart Enhance raises the dominant body size to at least this | 12 pt | 24 pt |
| Title, 28 pt | Smart Enhance title minimum; the title must also be at least 1.25x the (scaled) body size | 21 pt | 42 pt |
| Footer, 12 pt | Smart Enhance footer / slide-number minimum | 9 pt | 18 pt |
| Small text, 12 pt | `small-text` finding: any text run below this | 9 pt | 18 pt |
| Low body size, 14 pt | `low-body-size` finding: dominant body size below this | 10.5 pt | 21 pt |
| Large text, 18 pt | contrast classification: regular text from this size counts as large (3:1 needed instead of 4.5:1) | 13.5 pt | 27 pt |
| Large bold text, 14 pt | contrast classification: bold text from this size counts as large | 10.5 pt | 21 pt |

A subtitle has no threshold of its own: Smart Enhance only keeps it from being smaller than the (scaled) body size. A floor is a minimum, so a value already at or above it never changes, and Smart Enhance never lowers a size. On a 10 x 5.625 in slide a 10 pt footer is therefore left alone (the scaled floor is 9 pt), and on a 20 x 11.25 in slide a 21 pt body is raised to 24 pt. Where `s` is not 1, the Smart Enhance change reasons and the `small-text` / `low-body-size` messages say that the reference minimum was scaled; at `s = 1` no such note is added.

**Not scaled**: the 4.5:1 and 3:1 contrast ratios, the margins (a share of the slide; Smart Enhance floors stay 5 / 5 / 5 / 4 %), the 1.25x title-to-body ratio itself, the 4 % margin-consistency tolerance and the 110-word density limit.

**Known limitations of this step**: `inferTitle` still uses a fixed 28 pt rule, and the `inconsistent-title-size` finding still uses a fixed 4 pt tolerance. Neither is scaled with the slide size.

### Design version 4: saved designs are re-analysed

`DESIGN_VERSION` is 4. A template whose stored `designVersion` differs from the running engine is not trusted: the next time its Preview opens, it is analysed again from the stored file and the new design and `designVersion` are saved over the old ones. The template file, ID, name, dates (including `updatedAt`, so library order does not change) and the chosen mode are untouched. Library cards keep showing the stored analysis until that Preview is opened. If saving fails (for example the Firestore rules were not redeployed), the Preview still shows the fresh analysis and the saved record stays at its old version until the next successful save. Versions 1 to 3 migrate the same way. Visible effect on existing templates: Smart Enhance floors, the `small-text` / `low-body-size` findings and the large-text contrast class now follow the slide size, so a deck that is not 13.333 x 7.5 in or 4:3 may get different Smart Enhance changes and findings than under version 3. Design JSON files downloaded earlier carry `version: 3` and are not rewritten.

### What the analysis does not do (also shown in the UI)

- Classes are heuristics, not understanding; check the confidence. Layout-name hints only work for English names.
- PPTX has no font weights beyond bold/regular; "Semibold" etc. appear in the font name.
- Charts, tables, SmartArt and video are detected by presence, position and size only.
- Colour transforms other than lumMod/lumOff (tint, shade, satMod) are not applied; affected colours are flagged approximate. Group rotation/flips and picture pixel resolution are not evaluated. Contrast is not judged where text sits on a picture or gradient (it is reported as "not evaluated" instead). Colour-map overrides on slides and layouts (PowerPoint's dark background styles) are applied.
- Slides without a title placeholder use the largest text box as the title: in the top half, or lower when it is at least 1.5x the next size (disclosed in the limits list). Recurring template text (a brand name at the same place on 30% of the slides, typed page numbers) is never the title and is left out of word counts and margins; empty text boxes with no fill or outline are not counted as bleed.
- Two rules still use fixed values and do not scale with the slide size: title inference for a slide whose only text box is the title candidate needs a run of at least 28 pt (`inferTitle`), and the `inconsistent-title-size` finding compares titles against the template median with a 4 pt tolerance. On slides much smaller or larger than 13.333 x 7.5 in, expect these two to fire too late or too early (not measured).
- Verified on python-pptx and pptxgenjs files only. PowerPoint, Keynote and Google Slides exports are untested.

## PPTX support: what is and isn't claimed

Supported: standard `.pptx` (OOXML PresentationML) up to 50 MB and 200 slides.
Rejected with a clear message: `.ppt`, password-protected files, `.pptm`, `.potx`, `.ppsx`, non-Office ZIPs, corrupt XML.

The preview is **structural**: text, picture and object boxes at their real positions. It does not render fonts, colours, gradients, effects, charts, SmartArt or video, and the UI says so. Fixtures were produced with python-pptx; files exported from PowerPoint, Keynote and Google Slides have **not** been tested yet.

## Run locally

```bash
python3 -m http.server 8765      # or: npm start
# open http://localhost:8765
```

ES modules need an HTTP origin; opening `index.html` as a file will not work.

## Firebase setup (about 10 minutes)

1. Create a Firebase project. Add a **Web app** and copy its config into `js/config.js` (`firebaseConfig`). This config is public by design; security comes from the rules below.
2. **Authentication**: enable *Google* and *Anonymous* providers. Add your Vercel domain under *Authorized domains*.
3. **Firestore** and **Storage**: create both (production mode).
4. Deploy the rules:
   ```bash
   npm i -g firebase-tools && firebase login
   firebase use --add            # pick your project
   firebase deploy --only firestore:rules,storage
   ```
5. **Storage CORS** (required for previews, which download the file in the browser). Edit `firebase/cors.json` with your origins, then:
   ```bash
   gcloud storage buckets update gs://YOUR_BUCKET --cors-file=firebase/cors.json
   ```
6. Optional but recommended: enable **App Check** and set a Firebase budget alert.

Local-mode data is not migrated when you switch to Firebase (see `PROJECT_HANDOVER.md`).

**Upgrading from Phase 1:** redeploy the rules (`firebase deploy --only firestore:rules,storage`). They now allow the owner to update only `design`, `designVersion` and `mode` on a template, and `mode` on projects. Without this, analysis results cannot be saved (the preview still shows them).

## Deploy to Vercel

Import the repository, framework preset **Other**, no build command, output directory = root. `vercel.json` adds a Content-Security-Policy and security headers. Routing is hash-based, so no rewrites are needed.

## Project layout

```
index.html            entry
css/styles.css        design tokens + components
vendor/jszip.min.js   JSZip 3.10.1 (MIT), vendored; no CDN at runtime
js/app.js             boot, shell, hash router
js/config.js          Firebase config, limits
js/store/             index.js contract, local.js (IndexedDB), firebase.js
js/pptx/              parser.js, validate.js, design.js (Phase 2 analysis)
js/ui/                modal, toast, states, upload flow, preview, design-panel
js/views/             dashboard, templates, projects, project, signin
firebase/             firestore.rules, storage.rules, cors.json
tests/                fixtures, make_fixture*.py generators (Phases 1 to 2.5), make_fixture_pptxgenjs.mjs, e2e.mjs, analysis.mjs, harness.mjs
```

## Tests

See `TESTING.md`. Short version: `npm i && npx playwright install chromium && npm test` (108 e2e checks, needs Chromium) and `npm run test:analysis` (206 browser-free checks of the analysis layer, 237 when the real FOOD deck is supplied; needs only `npm i`).

The results recorded for Phase 2.5 Step 3 come from an authoring sandbox where the analysis suite ran through Chromium's DOMParser instead of jsdom. **Still pending**: a native jsdom run of `npm run test:analysis`, and the Firebase manual checklist in `TESTING.md`. Neither is claimed as passed.
