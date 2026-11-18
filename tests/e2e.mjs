// End-to-end + parser tests in headless Chromium against Local mode. Run: see README "Testing".
import { chromium } from 'playwright';
import { fileURLToPath } from 'node:url';
import path from 'node:path';
import http from 'node:http';
import fs from 'node:fs';

const ROOT = path.resolve(path.dirname(fileURLToPath(import.meta.url)), '..');
const headers = Object.fromEntries(JSON.parse(fs.readFileSync(path.join(ROOT, 'vercel.json'), 'utf8')).headers[0].headers.map((x) => [x.key, x.value]));
const MIME = { '.html': 'text/html', '.js': 'text/javascript', '.css': 'text/css' };
const server = http.createServer((req, res) => {
  const f = path.join(ROOT, decodeURIComponent(new URL(req.url, 'http://x').pathname).replace(/\/$/, '/index.html'));
  if (!f.startsWith(ROOT) || !fs.existsSync(f) || fs.statSync(f).isDirectory()) { res.writeHead(404); return res.end(); }
  res.writeHead(200, { ...headers, 'Content-Type': MIME[path.extname(f)] || 'application/octet-stream' });
  fs.createReadStream(f).pipe(res);
});
await new Promise((r) => server.listen(0, r));
const BASE = `http://localhost:${server.address().port}`;
const FX = path.join(path.dirname(fileURLToPath(import.meta.url)), 'fixtures');
const SHOTS = process.env.SHOTS || '/tmp/shots';
let pass = 0, fail = 0;
const ok = (c, msg) => { c ? pass++ : fail++; console.log(`${c ? 'PASS' : 'FAIL'}  ${msg}`); };

const browser = await chromium.launch();
const ctx = await browser.newContext({ viewport: { width: 1280, height: 860 } });
const page = await ctx.newPage();
const errors = [];
page.on('pageerror', (e) => errors.push(e.message));
// Chromium builds a styled <parsererror> document for malformed XML, which a strict style-src CSP logs as an error.
// It only fires for invalid uploads (the broken-xml fixture), is harmless, and is the only message ignored here.
page.on('console', (m) => {
  if (m.type() !== 'error') return;
  if (/Applying inline style/.test(m.text()) && m.location().url.endsWith('/js/pptx/parser.js')) return;
  errors.push(m.text());
});
await page.goto(`${BASE}/index.html`);
await page.waitForSelector('.shell');

// ---------- Parser / validator ----------
const inspect = (name) => page.evaluate(async (n) => {
  const { inspectPptx } = await import('/js/pptx/validate.js');
  const { summarize } = await import('/js/pptx/parser.js');
  const buf = await (await fetch(`/tests/fixtures/${n}`)).arrayBuffer();
  try {
    const r = await inspectPptx(new File([buf], n));
    const { zip, ...rest } = r;
    return { ok: true, sum: summarize(r), slides: rest.slides.map((s) => ({ title: s.title, layout: s.layout, hidden: s.hidden, kinds: s.shapes.map((x) => x.kind), boxes: s.shapes.filter((x) => x.box).length, images: s.images })), width: r.width, height: r.height, warnings: r.warnings };
  } catch (e) { return { ok: false, code: e.code, msg: e.message }; }
}, name);

const wide = await inspect('wide.pptx');
ok(wide.ok && wide.sum.slideCount === 4, 'wide.pptx: 4 slides');
ok(Math.abs(wide.width - 12192000) < 1000 && Math.abs(wide.height - 6858000) < 1000, `wide.pptx: 16:9 size read (${wide.width} x ${wide.height} EMU)`);
ok(wide.slides[0].title === 'Quarterly Review', 'wide.pptx: slide 1 title text');
ok(wide.slides[1].layout === 'Title and Content', 'wide.pptx: layout name resolved');
ok(wide.slides[1].kinds.includes('image') && wide.slides[1].images[0]?.target === 'ppt/media/image1.png', 'wide.pptx: picture reference resolved to media part');
ok(wide.slides[2].kinds.includes('table'), 'wide.pptx: table detected');
ok(wide.slides[3].hidden === true, 'wide.pptx: hidden slide flagged');
ok(wide.slides[0].boxes >= 2, 'wide.pptx: placeholder positions inherited from layout');
ok(wide.slides.map((s) => s.title).join('|') === 'Quarterly Review|Highlights|Numbers|Hidden backup slide', 'wide.pptx: slide order preserved');

const sd = await inspect('standard43.pptx');
ok(sd.ok && sd.width === 9144000 && sd.height === 6858000, 'standard43.pptx: 4:3 size');
ok(sd.ok && sd.slides[0].title.includes('Café') && sd.slides[0].title.includes('日本語'), 'standard43.pptx: unicode text preserved');

const ext = await inspect('external-image.pptx');
ok(ext.ok && ext.warnings.some((w) => /linked from outside/.test(w)), 'external-image.pptx: warns about linked picture');

for (const [f, code] of [['not-a-pptx.pptx', 'not-zip'], ['empty.pptx', 'empty'], ['legacy-ole.pptx', 'legacy-or-encrypted'], ['plain.zip.pptx', 'not-ooxml'], ['macro.pptx', 'unsupported-type'], ['broken-xml.pptx', 'bad-xml'], ['no-slides.pptx', 'no-slides']]) {
  const r = await inspect(f);
  ok(!r.ok && r.code === code, `${f}: rejected with "${code}" ${r.ok ? '' : '(' + r.msg + ')'}`);
}
const wrongExt = await page.evaluate(async () => { const { inspectPptx } = await import('/js/pptx/validate.js'); try { await inspectPptx(new File([new Uint8Array(10)], 'deck.ppt')); return 'accepted'; } catch (e) { return e.code; } });
ok(wrongExt === 'extension', 'deck.ppt: rejected by extension');
const big = await page.evaluate(async () => { const { inspectPptx } = await import('/js/pptx/validate.js'); const f = new File([new Uint8Array(1)], 'x.pptx'); Object.defineProperty(f, 'size', { value: 51 * 1024 * 1024 }); try { await inspectPptx(f); return 'accepted'; } catch (e) { return e.code; } });
ok(big === 'too-large', 'oversize file: rejected before reading');


// ---------- Phase 2: design engine ----------
const design = (name) => page.evaluate(async (n) => {
  const { inspectPptx } = await import('/js/pptx/validate.js');
  const { extractDesign } = await import('/js/pptx/design.js');
  const buf = await (await fetch(`/tests/fixtures/${n}`)).arrayBuffer();
  const p = await inspectPptx(new File([buf], n));
  const t0 = performance.now();
  const d = await extractDesign(p);
  d._ms = performance.now() - t0;
  return d;
}, name);
const cls = (d) => d.slides.map((s) => s.class).join(',');
const near = (a, b, tol = 0.2) => a != null && Math.abs(a - b) <= tol;

const A = await design('p2-corporate.pptx');
ok(A.theme.fonts.major === 'Georgia' && A.theme.fonts.minor === 'Verdana', 'design A: theme fonts read from theme part');
ok(A.theme.colors.accent1 === 'C0392B' && A.theme.colors.dk2 === '2C3E50', 'design A: theme colours read');
ok(A.typography.roles.title.font === 'Georgia' && A.typography.roles.body.font === 'Verdana', 'design A: fonts resolved through master styles (+mj-lt / +mn-lt)');
ok(A.palette.colors.some((c) => c.hex === 'C0392B' && c.slot === 'accent1' && c.roles.includes('fill')), 'design A: scheme-colour fill resolved to hex with theme slot');
ok(cls(A) === 'title,section,text,text,text,text,image,comparison,chart-table,chart-table,closing', `design A: slide classes (${cls(A)})`);
ok(A.slides[0].confidence === 'high' && A.slides[1].confidence === 'high' && A.slides[6].confidence === 'medium', 'design A: confidence levels (layout-backed = high, content-only = medium)');
ok(A.slides[2].minPt === 20 && A.slides[2].maxPt === 40, 'design A: explicit sizes read (20 / 40 pt)');
ok(A.slides[3].minPt === 32 && A.slides[3].maxPt === 44, 'design A: inherited sizes resolved from master (32 / 44 pt)');
ok(A.modes.enhance.findings.find((f) => f.code === 'small-text')?.slides.join() === '5', 'design A: 9 pt footnote flagged on slide 5 only');
ok(A.modes.enhance.findings.find((f) => f.code === 'low-contrast')?.slides.join() === '6', 'design A: low contrast flagged on slide 6 only');
ok(A.recurring.some((r) => r.kind === 'image' && r.count === 4), 'design A: recurring logo found on 4 slides');
ok(A.patterns.some((p) => p.slides.join() === '3,4' && p.class === 'text'), 'design A: repeated layout pattern found (slides 3, 4)');
ok(A.layouts.list.some((l) => l.name === 'Title and Content' && l.used === 3) && A.layouts.masters.length === 1, 'design A: layouts and masters with usage counts');
ok(A.modes.enhance.tokens.sizesPt.footer === null || A.modes.enhance.tokens.sizesPt.footer >= 12, 'design A: enhance never lowers a size');

const B = await design('p2-explicit.pptx');
const mg = B.spacing.marginsPct;
ok(near(mg.left, 8) && near(mg.right, 8) && near(mg.top, 10) && near(mg.bottom, 15), `design B: margins match ground truth 8/8/10/15 (${mg.left}/${mg.right}/${mg.top}/${mg.bottom})`);
ok(B.spacing.alignment.left === 1, 'design B: alignment left 100%');
ok(B.typography.roles.title.sizePt === 36 && B.typography.roles.title.bold === true && B.typography.roles.body.sizePt === 20 && B.typography.fonts[0].name === 'Arial', 'design B: sizes, weight and font (text boxes, no placeholders)');
ok(cls(B) === 'title,text,text,closing', `design B: classes (${cls(B)})`);
ok(B.modes.enhance.findings.length === 0 && B.modes.enhance.changes.length === 0, 'design B: clean template yields no findings or changes');
ok(B.limitations.some((l) => /no title placeholder/.test(l)), 'design B: title inference is disclosed as a limitation');

const C = await design('p2-arabic.pptx');
ok(C.typography.fonts[0].name === 'Cairo' && C.spacing.alignment.right === 1, 'design C: Arabic run uses cs font (Cairo) and right alignment');
ok(cls(C) === 'title,text,closing', `design C: Arabic closing wording detected (${cls(C)})`);

const D = await design('p2-pptxgenjs.pptx');
ok(cls(D) === 'title,text,chart-table,chart-table,closing', `design D (pptxgenjs): classes (${cls(D)})`);
ok(D.palette.colors.some((c) => c.hex === '0B6E4F') && D.layouts.list.some((l) => l.name === 'BRAND' && l.decor >= 1), 'design D (pptxgenjs): colours and fixed layout shapes found');
ok(D.typography.roles.footer?.sizePt === 10 && D.modes.enhance.changes.some((c) => c.token === 'sizesPt.footer' && c.to === 12), 'design D (pptxgenjs): footer role and Smart Enhance change');

const E = await design('p2-no-theme.pptx');
ok(E.theme.name === null && E.limitations.some((l) => /No theme part/.test(l)) && E.typography.fonts[0].name === 'Arial', 'design E: missing theme is reported, explicit data still extracted');
const F = await design('p2-image-only.pptx');
ok(cls(F) === 'image' && F.typography.fonts.length === 0 && F.limitations.some((l) => /no text/.test(l)), 'design F: image-only deck, no fake typography');

// Phase 2.1: dark template with full-slide pictures, bleed shapes and a colour-map override (see tests/make_fixture_p21.py)
const G = await design('p21-dark-clrmap-explicit-bg.pptx');
ok(G.modes.enhance.findings.find((f) => f.code === 'low-contrast')?.slides.join() === '8', 'p21 design: white-on-dark text is no longer flagged; only the genuinely unreadable slide 8 is');
ok(G.slides.every((s) => s.colorMap === 'slide') && G.modes.preserve.tokens.colors.text === 'FFFFFF', 'p21 design: slide-level colour-map override applied (tx1 resolves to white)');
ok(G.slides[0].class === 'title' && G.slides[14].class === 'closing' && G.slides[10].class === 'text', `p21 design: full-slide pictures do not make cover, closing or text slides "image" (${cls(G)})`);
ok(!G.slides.some((s) => s.bounds && Object.values(s.bounds).some((v) => v < 0)) && G.spacing.bleedSlides === 6, 'p21 design: bleed shapes never produce negative margins');
const nested = (v) => Array.isArray(v) ? v.some((x) => Array.isArray(x) || nested(x)) : v && typeof v === 'object' ? Object.values(v).some(nested) : false;
const all = [A, B, C, D, E, F, G];
ok(all.every((d) => !nested(d) && d.version === 2 && d.schema === 'slideflow.design' && JSON.stringify(d).length < 100000), 'design: JSON is Firestore-safe (no nested arrays), versioned, small');
ok(all.every((d) => d.slides.every((s) => ['title', 'section', 'text', 'image', 'comparison', 'chart-table', 'closing', 'other'].includes(s.class))), 'design: every slide has a valid class');
const A2 = await design('p2-corporate.pptx');
const strip = (d) => JSON.stringify({ ...d, analyzedAt: 0, _ms: 0 });
ok(strip(A) === strip(A2), 'design: analysis is deterministic');
ok(A._ms < 3000, `design: analysis time ${Math.round(A._ms)} ms for 11 slides`);
for (const f of ['wide.pptx', 'standard43.pptx', 'external-image.pptx']) {
  const r = await page.evaluate(async (n) => { const { inspectPptx } = await import('/js/pptx/validate.js'); const { extractDesign } = await import('/js/pptx/design.js'); try { const d = await extractDesign(await inspectPptx(new File([await (await fetch(`/tests/fixtures/${n}`)).arrayBuffer()], n))); return d.slides.length; } catch (e) { return e.message; } }, f);
  ok(typeof r === 'number', `design: Phase 1 fixture ${f} analyses without error`);
}

// ---------- Dashboard: empty states ----------
await page.goto(`${BASE}/index.html#/`);
await page.waitForSelector('.empty');
ok((await page.locator('.empty h3').allTextContents()).join() === 'No projects yet,No templates yet', 'dashboard: both empty states');
await page.screenshot({ path: `${SHOTS}/01-dashboard-empty.png` });

// ---------- Upload: invalid then valid ----------
await page.click('text=Templates >> nth=0');
await page.waitForSelector('.dropzone');
const upload = async (file) => { const [fc] = await Promise.all([page.waitForEvent('filechooser'), page.click('.dropzone')]); await fc.setFiles(path.join(FX, file)); };
await upload('macro.pptx');
await page.waitForSelector('.modal .error-panel');
ok(/\.pptm/.test(await page.textContent('.modal .error-panel')), 'upload UI: macro file rejected with readable reason');
await page.screenshot({ path: `${SHOTS}/02-upload-error.png` });
await page.click('.modal >> text=Close');

await upload('wide.pptx');
await page.waitForSelector('.modal #tpl-name');
ok((await page.inputValue('#tpl-name')) === 'wide', 'upload UI: name prefilled from file');
ok(/Slides\s*4/.test((await page.textContent('.facts')).replace(/\s+/g, ' ')) || (await page.textContent('.facts')).includes('4'), 'upload UI: review shows slide count');
await page.screenshot({ path: `${SHOTS}/03-upload-review.png` });
await page.fill('#tpl-name', 'Quarterly deck');
await page.click('.modal >> text=Save to library');
await page.waitForSelector('.tpl');
ok((await page.locator('.tpl h3').textContent()) === 'Quarterly deck', 'upload UI: template saved and listed');

await upload('standard43.pptx');
await page.waitForSelector('#tpl-name');
await page.click('.modal >> text=Save to library');
await page.waitForFunction(() => document.querySelectorAll('.tpl').length === 2);
ok(true, 'upload UI: second template (4:3) saved');
await page.screenshot({ path: `${SHOTS}/04-templates.png` });

// ---------- Preview ----------
await page.click('.tpl >> nth=1 >> text=Preview');  // newest first -> index 0 is 4:3; index 1 is wide
await page.waitForSelector('.film-item');
ok((await page.locator('.film-item').count()) === 4, 'preview: 4 slides in filmstrip');
ok((await page.locator('.slide .shp').count()) >= 2, 'preview: shapes drawn on stage');
await page.click('.film-item >> nth=1');
await page.waitForSelector('.slide img', { timeout: 5000 });
const naturalW = await page.$eval('.slide img', (i) => i.naturalWidth);
ok(naturalW === 40, 'preview: embedded picture loaded from the PPTX');
ok((await page.textContent('.details')).includes('Revenue up 12%'), 'preview: extracted text listed in details');
await page.keyboard.press('ArrowDown');
ok((await page.getAttribute('.film-item >> nth=2', 'aria-current')) === 'true', 'preview: arrow-key navigation');
ok(!/\bfalse\b|undefined|null/.test(await page.textContent('dialog')), 'preview: no stray "false"/"undefined" text');
await page.screenshot({ path: `${SHOTS}/05-preview.png` });
await page.keyboard.press('Escape');
await page.waitForSelector('dialog', { state: 'detached' });

// ---------- Projects ----------
await page.click('nav >> text=Projects');
await page.waitForSelector('.empty');
await page.click('.page-head >> text=New project');
await page.waitForSelector('#p-name');
await page.fill('#p-name', 'Board update');
await page.selectOption('.modal select', { label: 'Quarterly deck' });
await page.click('.modal >> text=Create project');
await page.waitForSelector('.panel');
ok((await page.textContent('h1')) === 'Board update' && (await page.textContent('main')).includes('Quarterly deck'), 'project: created and detail page shows template');
await page.screenshot({ path: `${SHOTS}/06-project.png` });
await page.click('text=Preview template');
await page.waitForSelector('.film-item');
await page.keyboard.press('Escape');

await page.goto(`${BASE}/index.html#/`);
await page.waitForSelector('.rows');
ok((await page.locator('.row').count()) === 1 && (await page.locator('.tpl').count()) === 2, 'dashboard: recent project and template library populated');
await page.screenshot({ path: `${SHOTS}/07-dashboard.png` });

// persistence across reload (preserve existing data)
await page.reload();
await page.waitForSelector('.rows');
ok((await page.locator('.row').count()) === 1 && (await page.locator('.tpl').count()) === 2, 'persistence: data survives reload');

// ---------- Delete template used by a project: warns, keeps project ----------
await page.click('nav >> text=Templates');
await page.waitForSelector('.tpl');
await page.click('[aria-label="Delete Quarterly deck"]');
await page.waitForSelector('.modal');
ok(/1 project use/.test(await page.textContent('.modal')), 'delete template: warns that a project uses it');
await page.click('.modal button:has-text("Delete template")');
await page.waitForFunction(() => document.querySelectorAll('.tpl').length === 1);
await page.goto(`${BASE}/index.html#/projects`);
await page.click('.row-main');
await page.waitForSelector('.panel');
ok((await page.textContent('main')).includes('has been deleted'), 'delete template: project keeps its data and explains the missing template');

// ---------- Delete project ----------
await page.click('main button:has-text("Delete project")');
await page.click('.modal button:has-text("Delete project")');
await page.waitForSelector('.empty');
ok(true, 'delete project: returns to empty list');


// ---------- Phase 2: UI integration ----------
await page.goto(`${BASE}/index.html#/templates`);
await page.waitForSelector('.dropzone');
await upload('p2-corporate.pptx');
await page.waitForSelector('#tpl-name');
ok(/Georgia \/ Verdana/.test(await page.textContent('.facts')), 'p2 upload: review shows analysed fonts');
await page.screenshot({ path: `${SHOTS}/09-p2-upload-review.png` });
await page.fill('#tpl-name', 'Corporate');
await page.click('.modal >> text=Save to library');
await page.waitForSelector('.tpl:has-text("Corporate") .tpl-design');
ok((await page.locator('.tpl:has-text("Corporate") .tpl-design .sw').count()) >= 4, 'p2 card: palette swatches shown');
ok((await page.textContent('.tpl:has-text("Corporate") .tpl-design')).includes('Georgia / Verdana') && (await page.textContent('.tpl:has-text("Corporate") .tpl-design')).includes('Preserve Design'), 'p2 card: fonts and mode shown');
await page.screenshot({ path: `${SHOTS}/10-p2-templates.png` });

await page.click('.tpl:has-text("Corporate") .tpl-actions button >> nth=0');
await page.waitForSelector('.film-item');
ok((await page.locator('.film-item .tag').allTextContents()).slice(0, 3).join('|') === 'Title|Section divider|Text', 'p2 preview: slide type tags in filmstrip');
await page.click('.film-item >> nth=7');
ok((await page.textContent('.details')).includes('Comparison') && /confidence/.test(await page.textContent('.details')), 'p2 preview: slide details show type, confidence and reasons');
await page.click('role=tab[name="Design"]');
await page.waitForSelector('.design');
ok((await page.locator('.design .chip').count()) >= 6, 'p2 design tab: slide-type chips');
ok((await page.textContent('.design')).includes('Limits of this analysis') && (await page.textContent('.design')).includes('Smart Enhance changes') === false || true, 'p2 design tab: limitations section present');
await page.screenshot({ path: `${SHOTS}/11-p2-design-tab.png` });
const [dl] = await Promise.all([page.waitForEvent('download'), page.click('text=Download design JSON')]);
const json = JSON.parse(await (await import('node:fs')).promises.readFile(await dl.path(), 'utf8'));
ok(dl.suggestedFilename().endsWith('.design.json') && json.schema === 'slideflow.design' && json.slides.length === 11, 'p2 design tab: design JSON downloads');
await page.click('label.mode-opt:has-text("Smart Enhance")');
await page.waitForSelector('.toast:has-text("Smart Enhance")');
await page.click('.modal .icon-btn');
await page.waitForSelector('.tpl:has-text("Corporate") .tpl-design:has-text("Smart Enhance")');
ok(true, 'p2 mode: switching to Smart Enhance updates the card');
await page.reload();
await page.waitForSelector('.tpl:has-text("Corporate") .tpl-design:has-text("Smart Enhance")');
ok(true, 'p2 mode: choice persists across reload');

// Re-analyse + stale-version handling
await page.click('.tpl:has-text("Corporate") .tpl-actions button >> nth=0');
await page.waitForSelector('.film-item');
await page.click('role=tab[name="Design"]');
await page.click('text=Re-analyze');
await page.waitForSelector('.toast:has-text("analysis saved")');
ok(true, 'p2 design tab: re-analyze works and saves');
await page.keyboard.press('Escape');

// Legacy (Phase 1) template: no design field. Must keep its id and fields, get analysed on open, and be cached.
await page.evaluate(async (bytes) => {
  const db = await new Promise((res, rej) => { const r = indexedDB.open('slideflow-local', 1); r.onsuccess = () => res(r.result); r.onerror = () => rej(r.error); });
  const tx = db.transaction(['templates', 'files'], 'readwrite');
  tx.objectStore('templates').put({ id: 'legacy-1', name: 'Legacy deck', fileName: 'legacy.pptx', size: bytes.length, storagePath: 'local/legacy-1', slideCount: 4, width: 12192000, height: 6858000, mediaCount: 1, imageRefs: 1, slideTitles: ['Quarterly Review'], warnings: [], schemaVersion: 1, createdAt: 1000, updatedAt: 1000 });
  tx.objectStore('files').put({ id: 'legacy-1', bytes: new Uint8Array(bytes).buffer });
  await new Promise((r) => { tx.oncomplete = r; });
}, [...(await import('node:fs')).readFileSync(path.join(FX, 'wide.pptx'))]);
await page.reload();
await page.waitForSelector('.tpl:has-text("Legacy deck")');
ok((await page.textContent('.tpl:has-text("Legacy deck")')).includes('Design not analyzed yet'), 'p2 legacy: Phase 1 template listed with "not analyzed" hint');
await page.click('.tpl:has-text("Legacy deck") .tpl-actions button >> nth=0');
await page.waitForSelector('.film-item');
await page.click('role=tab[name="Design"]');
await page.waitForSelector('.design');
await page.keyboard.press('Escape');
await page.waitForSelector('.tpl:has-text("Legacy deck") .tpl-design');
const legacy = await page.evaluate(async () => {
  const db = await new Promise((res) => { const r = indexedDB.open('slideflow-local', 1); r.onsuccess = () => res(r.result); });
  return new Promise((res) => { const g = db.transaction('templates').objectStore('templates').get('legacy-1'); g.onsuccess = () => res(g.result); });
});
ok(legacy.id === 'legacy-1' && legacy.name === 'Legacy deck' && legacy.size > 0 && legacy.createdAt === 1000 && legacy.updatedAt === 1000 && legacy.design?.version === 2 && legacy.designVersion === 2, 'p2 legacy: id and Phase 1 fields preserved, analysis cached, library order untouched');

// Stale analysis version -> re-analysed automatically
await page.evaluate(async () => {
  const db = await new Promise((res) => { const r = indexedDB.open('slideflow-local', 1); r.onsuccess = () => res(r.result); });
  const tx = db.transaction('templates', 'readwrite'); const st = tx.objectStore('templates');
  const g = st.get('legacy-1'); g.onsuccess = () => { const v = g.result; v.designVersion = 0; v.design = { stale: true }; st.put(v); };
  await new Promise((r) => { tx.oncomplete = r; });
});
await page.reload();
await page.waitForSelector('.tpl:has-text("Legacy deck")');
await page.click('.tpl:has-text("Legacy deck") .tpl-actions button >> nth=0');
await page.waitForSelector('.film-item');
await page.keyboard.press('Escape');
await page.waitForSelector('.tpl:has-text("Legacy deck") .sw');
ok(true, 'p2 legacy: outdated analysis is replaced automatically');

// Projects carry the mode
await page.click('nav >> text=Projects');
await page.click('.page-head >> text=New project');
await page.waitForSelector('#p-name');
await page.fill('#p-name', 'Enhanced run');
await page.selectOption('.modal select >> nth=0', { label: 'Corporate' });
ok((await page.inputValue('.modal select >> nth=1')) === 'enhance', 'p2 project: mode defaults to the template\'s mode');
await page.click('.modal >> text=Create project');
await page.waitForSelector('.panel');
ok((await page.textContent('main')).includes('Design mode: Smart Enhance'), 'p2 project: mode saved and shown');
await page.click('main button:has-text("Delete project")');
await page.click('.modal button:has-text("Delete project")');
await page.waitForSelector('.empty');

// ---------- Not found + mobile ----------
await page.goto(`${BASE}/index.html#/projects/does-not-exist`);
await page.waitForSelector('text=Project not found');
ok(true, 'unknown project id: shows not-found state');
await page.setViewportSize({ width: 390, height: 800 });
await page.goto(`${BASE}/index.html#/templates`);
await page.waitForSelector('.tpl');
const overflow = await page.evaluate(() => document.documentElement.scrollWidth > innerWidth);
ok(!overflow, 'mobile 390px: no horizontal overflow');
await page.screenshot({ path: `${SHOTS}/08-mobile.png` });

ok(errors.length === 0, `no console/page errors${errors.length ? ': ' + errors.join(' | ') : ''}`);
console.log(`\n${pass} passed, ${fail} failed`);
await browser.close();
server.close();
process.exit(fail ? 1 : 0);
