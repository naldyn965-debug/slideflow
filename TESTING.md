# Testing

## Automated (`npm test`)

Runs in headless Chromium against Local mode, served with the real `vercel.json` headers (so the CSP is enforced). 42 checks.
Setup: `npm i`, `npx playwright install chromium`. Fixtures are committed; regenerate with `pip install python-pptx && npm run fixtures`.

Covered: parser (slide count/order, 16:9 and 4:3 sizes, titles, layout names, placeholder position inheritance, picture to media resolution, tables, groups, hidden slides, Unicode, linked pictures), validator rejections (empty, text file, legacy/encrypted header, plain ZIP, macro-enabled, corrupt XML, no slides, wrong extension, oversize), upload flow, preview (filmstrip, shapes, embedded picture loads, arrow keys), project create/open/delete, template delete with in-use warning, persistence across reload, not-found state, no horizontal overflow at 390 px, no console errors.

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
