# Testing

## Automated (`npm test`)

Runs in headless Chromium against Local mode, served with the real `vercel.json` headers (so the CSP is enforced). 92 checks (42 Phase 1, 50 Phase 2).
Setup: `npm i`, `npx playwright install chromium`. Fixtures are committed; regenerate with `pip install python-pptx` then `npm run fixtures` (pptxgenjs fixture: `npm i pptxgenjs && node tests/make_fixture_pptxgenjs.mjs`).

Covered: parser (slide count/order, 16:9 and 4:3 sizes, titles, layout names, placeholder position inheritance, picture to media resolution, tables, groups, hidden slides, Unicode, linked pictures), validator rejections (empty, text file, legacy/encrypted header, plain ZIP, macro-enabled, corrupt XML, no slides, wrong extension, oversize), upload flow, preview (filmstrip, shapes, embedded picture loads, arrow keys), project create/open/delete, template delete with in-use warning, persistence across reload, not-found state, no horizontal overflow at 390 px, no console errors.

Phase 2 coverage: ground-truth checks on six templates (python-pptx corporate 4:3 with edited theme, explicit-geometry 16:9, Arabic RTL, pptxgenjs, theme removed, picture-only): theme colours/fonts, style inheritance through master, scheme-colour fills, exact margins, alignment, class sequence and confidence, repeated patterns, recurring logo, small-text and low-contrast findings, Smart Enhance changes, limitations shown for missing data, Firestore-safe JSON, determinism; UI: upload review, card swatches, filmstrip tags, Design tab, JSON download, mode change and persistence, re-analyze, Phase 1 template upgraded in place (ID and fields kept), stale analysis refreshed, project mode.

## Manual checklist (needs a real Firebase project; not run in this build)

Sign-in
- [ ] Google sign-in works; guest sign-in works; sign-out returns to the sign-in screen
- [ ] Second account cannot see the first account's templates or projects

Templates
- [ ] Upload a real deck exported from PowerPoint, Keynote and Google Slides; compare slide count and order with the source
- [ ] Upload shows progress; after saving, the file exists at `users/{uid}/templates/{id}.pptx`
- [ ] Preview opens (if it fails with a network error, fix bucket CORS)
- [ ] Delete removes both the Storage file and the Firestore document
- [ ] Upload with the network cut mid-way: no Firestore record is left behind

Rules (Firebase emulator or console Rules Playground)
- [ ] Unauthenticated read/write denied; other user's path denied
- [ ] Storage rejects non-PPTX content type, files over 50 MB, and overwrites
- [ ] Firestore rejects extra fields, `slideCount` over 200, and a `storagePath` that doesn't match the path

Deployment
- [ ] No CSP violations in the console on the Vercel URL while signing in, uploading and previewing
- [ ] Mobile (iOS Safari, Android Chrome): upload, preview and create project

Large files
- [ ] 40 MB / 150-slide deck: parse time acceptable, UI stays responsive

Phase 2 additions
- [ ] Analyse real decks from PowerPoint, Keynote and Google Slides; check slide classes against your own judgement and note misclassifications
- [ ] A Phase 1 template in Firestore opens, analyses and saves (needs the updated rules deployed); the template keeps its ID and file
- [ ] Firestore rules (emulator): owner can update only `design`, `designVersion`, `mode`; changing `name` or `size` on update is denied
- [ ] 150+ slide deck: analysis time acceptable
- [ ] Deck with gradient or picture backgrounds: limitations list mentions contrast not evaluated
