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
