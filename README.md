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
  - **Smart Enhance**: same fonts, palette and layouts, but sizes, margins and contrast are lifted to readable minimums. Findings (small text, low contrast, too many fonts, inconsistent margins...) and the exact list of changes are shown.
  Modes are a stored contract for Phase 3 slide generation. **Nothing modifies your PPTX file.**
- **Existing templates** (uploaded in Phase 1) keep their IDs and fields. They show "Design not analyzed yet"; opening Preview analyses them once and caches the result. Outdated analyses are replaced the same way.
- **Download design JSON** from the Design tab.

### What the analysis does not do (also shown in the UI)

- Classes are heuristics, not understanding; check the confidence. Layout-name hints only work for English names.
- PPTX has no font weights beyond bold/regular; "Semibold" etc. appear in the font name.
- Charts, tables, SmartArt and video are detected by presence, position and size only.
- Colour transforms other than lumMod/lumOff (tint, shade, satMod) are not applied; affected colours are flagged approximate. Group rotation/flips and picture pixel resolution are not evaluated. Contrast is not judged where text sits on a picture or gradient (it is reported as "not evaluated" instead). Colour-map overrides on slides and layouts (PowerPoint's dark background styles) are applied.
- Slides without a title placeholder use the top, largest text box as the title (disclosed in the limits list).
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
tests/                fixtures, make_fixtures.py, make_fixtures_p2.py, make_fixture_pptxgenjs.mjs, e2e.mjs
```

## Tests

See `TESTING.md`. Short version: `npm i && npx playwright install chromium && npm test` (e2e, needs Chromium) and `npm run test:analysis` (55 browser-free checks of the analysis layer, needs only `npm i`).
