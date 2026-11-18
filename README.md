# SlideFlow, Phase 1

Dashboard, PPTX template management and a PPTX parser. Plain HTML, CSS and ES modules. No build step, no framework.
Out of scope for this phase: AI generation, research, slide editor, export.

## What works

- **Dashboard**: sidebar, recent projects, template library, create project, loading / empty / error states.
- **Templates**: upload (click or drag and drop), validate, review, save, list, structural preview, delete.
- **Projects**: create (optionally from a template), list, open, delete. A project is a record pointing at a template; the editor comes later.
- **Parser** (`js/pptx/parser.js`): slide count and order, slide size, per-slide text, shape positions, layout names, hidden slides, image references (with the real picture shown in the preview when the browser supports the format).
- **Two storage modes**, same code path for the UI:
  - **Local mode** (default, no setup): IndexedDB in this browser.
  - **Firebase mode** (Auth + Firestore + Storage): switches on when `js/config.js` has a Firebase config.

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
js/pptx/              parser.js, validate.js
js/ui/                modal, toast, states, upload flow, preview
js/views/             dashboard, templates, projects, project, signin
firebase/             firestore.rules, storage.rules, cors.json
tests/                fixtures, make_fixtures.py, e2e.mjs
```

## Tests

See `TESTING.md`. Short version: `npm i && npx playwright install chromium && npm test`.
