// Phase 3.2.1 tests: element selection in the Rendered tab.  Run: node tests/selection.mjs
// Same setup as tests/render.mjs: headless Chromium (playwright) against the app served with the real vercel.json headers, so CSP violations fail the run.
// Fixture: tests/fixtures/p321-selection.pptx (tests/make_fixture_p321.py). Its coordinates are in POINTS and are hard-coded below; the slide is drawn
// 960 px wide at the top-left of the page, so 1 pt = 1 px and a click at (x, y) is a click at that slide point.
import { chromium } from 'playwright';
import http from 'node:http';
import { readFileSync, readdirSync, existsSync, statSync } from 'node:fs';
import path from 'node:path';
import { fileURLToPath } from 'node:url';

const ROOT = path.resolve(path.dirname(fileURLToPath(import.meta.url)), '..');
const MIME = { html: 'text/html; charset=utf-8', js: 'text/javascript; charset=utf-8', css: 'text/css; charset=utf-8', json: 'application/json', svg: 'image/svg+xml', png: 'image/png', ttf: 'font/ttf', woff2: 'font/woff2', txt: 'text/plain; charset=utf-8', pptx: 'application/octet-stream' };
const vercel = JSON.parse(readFileSync(path.join(ROOT, 'vercel.json'), 'utf8'));
const headers = Object.fromEntries((vercel.headers?.[0]?.headers || []).map((h) => [h.key, h.value]));
const server = http.createServer((req, res) => {
  const u = decodeURIComponent(req.url.split('?')[0]);
  const f = path.join(ROOT, u === '/' ? 'index.html' : u);
  if (!f.startsWith(ROOT) || !existsSync(f) || statSync(f).isDirectory()) { res.writeHead(404); res.end('nf'); return; }
  res.writeHead(200, { ...headers, 'Content-Type': MIME[f.split('.').pop()] || 'application/octet-stream' });
  res.end(readFileSync(f));
});
await new Promise((r) => server.listen(0, r));
const BASE = `http://localhost:${server.address().port}`;

let pass = 0; let fail = 0;
const ok = (c, msg, extra) => { if (c) { pass++; console.log('  ok   ' + msg); } else { fail++; console.log('FAIL   ' + msg + (extra !== undefined ? '  -> ' + JSON.stringify(extra).slice(0, 500) : '')); } };
const section = (t) => console.log('\n# ' + t);
const FX = path.join(ROOT, 'tests/fixtures');
const REAL = path.join(ROOT, 'tests/real');
const SEL = path.join(FX, 'p321-selection.pptx');
const b64 = (f) => readFileSync(f).toString('base64');
const realFiles = [...(existsSync(REAL) ? readdirSync(REAL).filter((f) => f.endsWith('.pptx')).map((f) => path.join(REAL, f)) : []), ...(process.env.REAL_PPTX || '').split(':').filter(Boolean)];

const browser = await chromium.launch();
const ctx = await browser.newContext({ viewport: { width: 1400, height: 900 } });
const page = await ctx.newPage();
const consoleErrors = [];
const watch = (pg, tag) => {
  pg.on('console', (m) => { if (m.type() !== 'error') return; if (/Applying inline style/.test(m.text()) && m.location().url.endsWith('/js/pptx/parser.js')) return; consoleErrors.push(tag + m.text()); });
  pg.on('pageerror', (e) => consoleErrors.push(tag + 'pageerror: ' + e.message));
};
watch(page, '');
await page.goto(BASE + '/index.html');
await page.waitForFunction(() => !!globalThis.JSZip);

// page-side helpers: __load (deck -> parsed + scene module), __draw (render one slide at 960 px and attach selection)
await page.evaluate(() => {
  globalThis.__load = async (b64) => {
    const bytes = Uint8Array.from(atob(b64), (c) => c.charCodeAt(0));
    const parser = await import('/js/pptx/parser.js');
    const parsed = await parser.parsePptx(bytes);
    const scene = await import('/js/model/scene.js');
    return { parsed, parser, scene };
  };
  const h = document.createElement('div'); h.id = 'sf-host'; document.body.append(h);
  const st = h.style; st.position = 'fixed'; st.left = '0'; st.top = '0'; st.width = '960px'; st.zIndex = '99999'; st.background = '#888';
  globalThis.__draw = async (b64, idx, shared) => {
    const { parsed, scene } = await globalThis.__load(b64);
    const rn = await import('/js/render/render.js'); const sl = await import('/js/render/selection.js');
    const sc = await scene.buildScene(parsed, idx);
    globalThis.__s?.sel.dispose();
    const hnd = await rn.renderSlide(sc, document.getElementById('sf-host'), { width: 960 });
    const changes = []; const moves = []; const resizes = [];
    const sel = sl.attachSelection(hnd, sc, { onChange: (id) => changes.push(id), onMove: (m) => moves.push(m.phase), onResize: (m) => resizes.push(m.phase), edits: shared ? globalThis.__E : undefined }); // shared: the Phase 3.2.3-B edit-state (globalThis.__E), kept across redraws
    globalThis.__s = { hnd, sc, sel, parsed, changes, moves, resizes };
    return { slideWidth: sc.width, ids: sc.elements.map((e) => e.id) };
  };
});
const draw = (file, idx, shared = false) => page.evaluate(({ b, idx, shared }) => globalThis.__draw(b, idx, shared), { b: b64(file), idx, shared });
/** [x, y, pointerType?] list -> element names (or null) that selection would pick there */
const picks = (pts) => page.evaluate((pts) => { const s = globalThis.__s; return pts.map(([x, y, t]) => { const id = s.sel.pick(x, y, t || 'mouse'); return id ? s.sc.byId.get(id).name : null; }); }, pts);
const table = async (label, rows) => {
  const got = await picks(rows.map((r) => r.slice(0, 3).filter((v, i) => i < 2 || v)));
  rows.forEach((r, k) => ok(got[k] === r[3], `${label}: ${r[4]} -> ${r[3] ?? 'nothing (clears)'}`, { at: r.slice(0, 3), got: got[k] }));
};
// row = [x, y, pointerType|'', expected name|null, description]

// ======================================================================================================
section('hit-testing on known geometry (p321-selection, slide 1; coordinates in points)');
const ids1 = await draw(SEL, 0);
ok(ids1.slideWidth === 12192000, 'the fixture slide is 960 pt wide, so 1 pt = 1 px here');
await table('overlap', [
  [100, 100, '', 'back-rect', 'back-rect only'],
  [250, 100, '', 'back-rect', 'visible part of back-rect beside the overlap'],
  [250, 200, '', 'front-rect', 'overlap of back-rect and front-rect: the later (visually higher) element'],
  [450, 300, '', 'front-rect', 'front-rect only'],
]);
await table('non-rectangular geometry', [
  [620, 160, '', 'circle', 'centre of the circle'],
  [530, 70, '', 'under-rect', 'corner of the circle\'s frame (outside the circle) falls through to the rect underneath'],
  [750, 160, '', 'under-rect', 'rect beside the circle'],
  [830, 240, '', 'pic', 'centre of the ellipse-cropped picture'],
  [783, 193, '', null, 'corner of the picture frame outside its ellipse crop'],
]);
await table('rotation', [
  [860, 100, '', 'rot-rect', 'centre of the rotated rect'],
  [902, 142, '', 'rot-rect', 'on the rotated rect but outside its unrotated frame'],
  [785, 75, '', null, 'inside the unrotated frame but outside the rotated rect'],
  [895, 65, '', null, 'inside the rotated rect\'s bounding box but outside the rect'],
]);
await table('outline-only and nested', [
  [740, 400, '', 'inner-fill', 'inner-fill is drawn above the outline-only shape'],
  [660, 340, '', null, 'inside an outline-only shape (no fill): falls through to empty space'],
  [760, 320, '', 'outline-only', 'on the 3 pt outline'],
]);
await table('text and table', [
  [170, 320, '', 'textbox', 'text frame beyond the glyphs (no fill, no outline)'],
  [250, 470, '', 'table', 'a table cell'],
]);
await table('groups', [
  [400, 40, '', 'g-child-a', 'child of a group is selected, not the group'],
  [455, 120, '', 'g-child-b', 'second child of the same group'],
  [485, 25, '', null, 'inside the group\'s frame but over none of its children'],
]);
await table('not selectable', [
  [110, 500, '', null, 'hidden element'],
  [500, 525, '', null, 'master decoration (the strip is drawn, but inert)'],
  [30, 30, '', null, 'empty slide area over the full-slide layout decoration'],
  [950, 300, '', null, 'empty slide area at the edge'],
]);
await table('hairlines (1 pt line, 2 pt diagonal)', [
  [200, 400, '', 'h-line', 'exactly on the 1 pt line'],
  [200, 402, '', 'h-line', '2 px off the 1 pt line, mouse'],
  [200, 407, '', null, '7 px off the line, mouse: out of reach'],
  [200, 407, 'touch', 'h-line', '7 px off the line, finger: within reach'],
  [200, 416, 'touch', null, '16 px off the line, finger: out of reach'],
  [370, 400, 'touch', 'h-line', '10 px past the line\'s end, finger: round end reach'],
  [500, 440, '', 'diag-line', 'on the diagonal'],
  [590, 385, '', null, 'corner of the diagonal\'s bounding box, 93 pt from the line (a bounding box must not steal clicks)'],
  [497.4, 444.3, '', null, '5 px beside the diagonal, mouse'],
  [496.4, 446.0, 'touch', 'diag-line', '7 px beside the diagonal, finger'],
]);

section('stacking, flips and rotation (p321-selection, slide 3)');
await draw(SEL, 2);
await table('slide 3', [
  [200, 200, '', 'dup-top', 'two elements with the same frame: the later one is on top'],
  [650, 137.5, '', 'flip-line', 'on the horizontally flipped line'],
  [550, 137.5, '', null, 'on the mirror image (the unflipped diagonal): nothing'],
  [250, 420, '', 'rot-line', 'centre of the rotated line'],
  [314.95, 457.5, '', 'rot-line', 'on the rotated line, off the horizontal'],
  [314.95, 420, '', null, 'where the unrotated line would be'],
  [550, 432, '', 'card-over', 'a hairline 2 px away loses to the filled element that is drawn above it (exactly under the pointer)'],
  [450, 492, '', 'hair-over', 'a hairline 2 px away wins over the filled element underneath it (it is drawn above)'],
  [450, 497, '', 'card-under', '7 px from that hairline, mouse: out of reach, so the element under the pointer'],
  [450, 497, 'touch', 'hair-over', '7 px from that hairline, finger: within reach and above'],
]);

section('select() / clear() / ids (programmatic API, the same code path as a click)');
await draw(SEL, 0);
const api = await page.evaluate(() => {
  const { sel, sc, hnd, changes } = globalThis.__s; const out = {};
  const front = sc.elements.find((e) => e.name === 'front-rect'); const back = sc.elements.find((e) => e.name === 'back-rect');
  const grp = sc.elements.find((e) => e.name === 'grp'); const hid = sc.elements.find((e) => e.name === 'hidden-rect'); const wash = [...sc.byId.values()].find((e) => e.name === 'layout-wash');
  out.ids = { front: front.id, back: back.id, wash: wash.id };
  out.startNone = sel.id === null && !hnd.root.querySelector('.sf-sel');
  out.selFront = sel.select(front.id); out.afterFront = sel.id; out.overlays1 = hnd.root.querySelectorAll('.sf-sel').length; out.forAttr = hnd.root.querySelector('.sf-sel')?.dataset.for;
  out.again = sel.select(front.id); out.changesAfterAgain = changes.length; // selecting the selected element is not a change
  out.selBack = sel.select(back.id); out.afterBack = sel.id; out.overlays2 = hnd.root.querySelectorAll('.sf-sel').length;
  out.refuseGroup = sel.select(grp.id); out.refuseHidden = sel.select(hid.id); out.refuseLayout = sel.select(wash.id); out.refuseUnknown = sel.select('nope#1');
  out.stillBack = sel.id;
  sel.clear(); out.cleared = sel.id === null && !hnd.root.querySelector('.sf-sel'); sel.clear();
  out.changes = changes.slice();
  return out;
});
ok(api.ids.front === 'ppt/slides/slide1.xml#3' && api.ids.back === 'ppt/slides/slide1.xml#2', 'ids are the stable Scene Model ids (part#cNvPr id)', api.ids);
ok(api.startNone, 'nothing is selected and no outline exists before the first selection');
ok(api.selFront && api.afterFront === api.ids.front && api.overlays1 === 1 && api.forAttr === api.ids.front, 'select(id) selects that scene id and draws exactly one outline for it', api);
ok(api.again === true && api.changesAfterAgain === 1, 'selecting the already selected element is not a change');
ok(api.selBack && api.afterBack === api.ids.back && api.overlays2 === 1, 'selecting another element replaces the selection and the outline', api);
ok(!api.refuseGroup && !api.refuseHidden && !api.refuseLayout && !api.refuseUnknown && api.stillBack === api.ids.back, 'groups, hidden elements, layout elements and unknown ids cannot be selected (and leave the selection alone)', api);
ok(api.cleared, 'clear() removes the selection and its outline; clearing twice is harmless');
ok(JSON.stringify(api.changes) === JSON.stringify([api.ids.front, api.ids.back, null]), 'onChange reports each change once: front, back, cleared', api.changes);

// ======================================================================================================
section('pointer protocol (synthetic pointer events: the edges real input cannot produce on demand)');
await draw(SEL, 0);
const proto = await page.evaluate(() => {
  const { sel, sc, hnd } = globalThis.__s; const root = hnd.root; const R = root.getBoundingClientRect(); const out = {};
  const ev = (type, x, y, o = {}) => root.dispatchEvent(new PointerEvent(type, { bubbles: true, cancelable: true, composed: true, pointerType: 'touch', isPrimary: true, pointerId: 1, button: 0, clientX: R.left + x, clientY: R.top + y, ...o }));
  const name = () => (sel.id ? sc.byId.get(sel.id).name : null);
  const tap = (x1, y1, x2, y2, o1 = {}, o2 = {}) => { sel.clear(); ev('pointerdown', x1, y1, o1); ev('pointerup', x2 ?? x1, y2 ?? y1, o2); return name(); };
  out.touchTap = tap(450, 300);
  out.touchWithin = tap(450, 300, 458, 300); // 8 px, touch limit 10
  out.touchBeyond = tap(450, 300, 465, 300); // 15 px
  out.mouseWithin = tap(450, 300, 453, 300, { pointerType: 'mouse' }, { pointerType: 'mouse' });
  out.mouseBeyond = tap(450, 300, 456, 300, { pointerType: 'mouse' }, { pointerType: 'mouse' });
  out.pen = tap(450, 300, 450, 300, { pointerType: 'pen' }, { pointerType: 'pen' });
  sel.clear(); ev('pointerdown', 450, 300); ev('pointercancel', 450, 300); ev('pointerup', 450, 300); out.cancelled = name();
  sel.clear(); ev('pointerdown', 450, 300); ev('pointerup', 450, 300, { pointerId: 2 }); out.otherPointer = name();
  sel.clear(); ev('pointerdown', 450, 300, { isPrimary: false, pointerId: 2 }); ev('pointerup', 450, 300, { isPrimary: false, pointerId: 2 }); out.secondary = name();
  sel.clear(); ev('pointerdown', 450, 300); ev('pointerdown', 100, 100, { isPrimary: false, pointerId: 2 }); ev('pointerup', 450, 300); out.secondFingerDuring = name();
  sel.clear(); ev('pointerup', 450, 300); out.upWithoutDown = name();
  sel.clear(); ev('pointerdown', 450, 300, { pointerType: 'mouse', button: 2 }); ev('pointerup', 450, 300, { pointerType: 'mouse', button: 2 }); out.rightButton = name();
  sel.clear(); ev('pointerdown', 450, 300, { pointerType: 'mouse', button: 1 }); ev('pointerup', 450, 300, { pointerType: 'mouse', button: 1 }); out.middleButton = name();
  // the press position decides: pressed on front-rect, released 3 px away over empty space is still a click on front-rect
  out.pressDecides = tap(498, 300, 502, 300, { pointerType: 'mouse' }, { pointerType: 'mouse' });
  // a press that starts on empty space and ends within the travel limit clears
  sel.select(sc.elements.find((e) => e.name === 'front-rect').id); ev('pointerdown', 30, 30); ev('pointerup', 31, 30); out.emptyClears = name();
  // a tab that is not showing (width 0): picking is safe and finds nothing
  const host = document.getElementById('sf-host'); host.style.display = 'none';
  try { out.hiddenPick = sel.pick(450, 300); } catch (e) { out.hiddenPick = 'threw ' + e; }
  host.style.display = '';
  // after dispose nothing reacts and no outline is left
  sel.select(sc.elements.find((e) => e.name === 'front-rect').id); sel.dispose();
  out.afterDispose = { outline: root.querySelectorAll('.sf-sel').length, id: sel.id };
  ev('pointerdown', 100, 100); ev('pointerup', 100, 100); out.afterDisposeClick = root.querySelectorAll('.sf-sel').length;
  return out;
});
ok(proto.touchTap === 'front-rect' && proto.pen === 'front-rect', 'a press and release of the same pointer selects (touch and pen)', proto);
ok(proto.touchWithin === 'front-rect' && proto.touchBeyond === null, 'touch may travel up to 10 px and still be a tap; 15 px is not', proto);
ok(proto.mouseWithin === 'front-rect' && proto.mouseBeyond === null, 'mouse may travel up to 4 px and still be a click; 6 px is a drag', proto);
ok(proto.cancelled === null, 'pointercancel (the browser took the touch for scrolling or zooming) ends the tap without a selection', proto);
ok(proto.otherPointer === null && proto.secondary === null && proto.secondFingerDuring === null && proto.upWithoutDown === null, 'a release by another pointer, a non-primary pointer, a second finger during the press, or a release without a press selects nothing', proto);
ok(proto.rightButton === null && proto.middleButton === null, 'right and middle mouse buttons do not select', proto);
ok(proto.pressDecides === 'front-rect', 'the element is the one under the pointer when it was pressed', proto);
ok(proto.emptyClears === null, 'a tap on empty space clears an existing selection', proto);
ok(proto.hiddenPick === null, 'picking on a slide that is not showing (display: none) is safe and finds nothing', proto);
ok(proto.afterDispose.outline === 0 && proto.afterDispose.id === null && proto.afterDisposeClick === 0, 'dispose() removes the outline and the listeners', proto);

// ======================================================================================================
section('outline accuracy: the outline sits exactly on the element\'s frame (rotation, flips, nested groups; expected corners computed independently by tests/make_fixture_p31.py)');
{
  const exp = JSON.parse(readFileSync(path.join(FX, 'p31-geometry.expected.json'), 'utf8')).elements;
  await draw(path.join(FX, 'p31-geometry.pptx'), 0);
  const res = await page.evaluate((exp) => {
    const { sel, sc, hnd } = globalThis.__s; const out = [];
    const u = 960 / (sc.width / 12700);
    for (const [name, e] of Object.entries(exp)) {
      if (e.kind === 'group') continue;
      const el = sc.elements.find((x) => x.name === name);
      if (!el || !sel.select(el.id)) { out.push({ name, missing: true }); continue; }
      const o = hnd.root.querySelector('.sf-sel'); const r = o.getBoundingClientRect(); const root = hnd.root.getBoundingClientRect();
      const xs = e.corners.filter((_, i) => i % 2 === 0).map((v) => (v / 12700) * u); const ys = e.corners.filter((_, i) => i % 2 === 1).map((v) => (v / 12700) * u);
      const box = { l: Math.min(...xs), r: Math.max(...xs), t: Math.min(...ys), b: Math.max(...ys) };
      out.push({ name, w: box.r - box.l, h: box.b - box.t, dl: r.left - root.left - box.l, dr: r.right - root.left - box.r, dt: r.top - root.top - box.t, db: r.bottom - root.top - box.b, shear: !!el.world.shear });
    }
    return out;
  }, exp);
  const meas = res.filter((r) => !r.missing && r.w >= 6 && r.h >= 6 && !r.shear);
  ok(res.length >= 10 && res.every((r) => !r.missing), `all ${res.length} non-group elements of p31-geometry can be selected`, res.filter((r) => r.missing));
  const worst = Math.max(...meas.map((r) => Math.max(Math.abs(r.dl), Math.abs(r.dr), Math.abs(r.dt), Math.abs(r.db))));
  ok(meas.length >= 10 && worst <= 1, `outline bounding box matches the independently computed frame within 1 px for ${meas.length} elements (worst ${worst.toFixed(2)} px)`, meas.filter((r) => Math.abs(r.dl) > 1 || Math.abs(r.dr) > 1 || Math.abs(r.dt) > 1 || Math.abs(r.db) > 1));
  ok(res.some((r) => /rot|flip/.test(r.name)) && meas.some((r) => /g08|g07|g06/.test(r.name)), 'the checked set includes rotated, flipped and group-nested elements');
}
{
  await draw(SEL, 0);
  const r = await page.evaluate(() => {
    const { sel, sc, hnd } = globalThis.__s; const el = sc.elements.find((e) => e.name === 'h-line'); sel.select(el.id);
    const o = hnd.root.querySelector('.sf-sel').getBoundingClientRect();
    return { l: o.left, r: o.right, cy: (o.top + o.bottom) / 2, h: o.height };
  });
  ok(Math.abs(r.l - 60) <= 1 && Math.abs(r.r - 360) <= 1 && Math.abs(r.cy - 400) <= 1 && r.h >= 6 - 0.5, 'a zero-height line gets a visible outline (>= 6 px) kept centred on the line, same extent along it', r);
  const edge = await page.evaluate(() => {
    const { sel, sc, hnd } = globalThis.__s; sel.select(sc.elements.find((e) => e.name === 'diag-line').id);
    const cs = getComputedStyle(hnd.root.querySelector('.sf-sel'));
    return { pe: cs.pointerEvents, border: cs.borderTopWidth, style: cs.borderTopStyle, pos: cs.position, box: cs.boxSizing };
  });
  ok(edge.pe === 'none' && edge.border === '2px' && edge.style === 'solid' && edge.box === 'border-box', 'the outline is a 2 px border inside the frame (visible at the slide edge) and never takes pointer events', edge);
}

// ======================================================================================================
section('resize handles (Phase 3.2.2-A): 8 visual handles on the outline, positioned exactly, no behaviour');
{
  const exp = JSON.parse(readFileSync(path.join(FX, 'p31-geometry.expected.json'), 'utf8')).elements;
  await draw(path.join(FX, 'p31-geometry.pptx'), 0);
  const res = await page.evaluate((exp) => {
    const { sel, sc, hnd } = globalThis.__s; const out = []; const u = 960 / (sc.width / 12700);
    const none = hnd.root.querySelectorAll('.sf-hdl').length;
    for (const [name, e] of Object.entries(exp)) {
      if (e.kind === 'group') continue;
      const el = sc.elements.find((x) => x.name === name); sel.select(el.id);
      const hs = [...hnd.root.querySelectorAll('.sf-sel .sf-hdl')]; const R = hnd.root.getBoundingClientRect();
      const centres = hs.map((h) => { const r = h.getBoundingClientRect(); return [r.left + r.width / 2 - R.left, r.top + r.height / 2 - R.top]; });
      const c = []; for (let i = 0; i < 8; i += 2) c.push([(e.corners[i] / 12700) * u, (e.corners[i + 1] / 12700) * u]);
      const mid = (a, b) => [(a[0] + b[0]) / 2, (a[1] + b[1]) / 2];
      const want = [c[0], c[1], c[2], c[3], mid(c[0], c[1]), mid(c[1], c[2]), mid(c[2], c[3]), mid(c[3], c[0])]; // as a SET: a flip swaps which corner is which
      const used = new Set(); let worst = 0;
      for (const w of want) { let bi = -1; let bd = 1e9; centres.forEach((p, i) => { const d = Math.hypot(p[0] - w[0], p[1] - w[1]); if (!used.has(i) && d < bd) { bd = d; bi = i; } }); used.add(bi); worst = Math.max(worst, bd); }
      const wpx = Math.hypot(c[1][0] - c[0][0], c[1][1] - c[0][1]); const hpx = Math.hypot(c[2][0] - c[1][0], c[2][1] - c[1][1]);
      out.push({ name, n: hs.length, kinds: hs.map((h) => h.dataset.h).sort().join(), worst, big: wpx >= 40 && hpx >= 40, shear: !!el.world.shear, rot: el.world.rot });
    }
    sel.clear();
    return { out, none, after: hnd.root.querySelectorAll('.sf-hdl').length };
  }, exp);
  const KINDS = ['e', 'n', 'ne', 'nw', 's', 'se', 'sw', 'w'].join();
  ok(res.none === 0 && res.after === 0, 'no handles exist before a selection and none are left after clear()');
  ok(res.out.every((r) => r.n === 8 && r.kinds === KINDS), `every selected element gets exactly 8 handles: nw n ne e se s sw w (${res.out.length} elements)`, res.out.filter((r) => r.n !== 8));
  const meas = res.out.filter((r) => r.big && !r.shear);
  const worst = Math.max(...meas.map((r) => r.worst));
  ok(meas.length >= 8 && worst <= 1, `handle centres sit on the independently computed corners and edge midpoints within 1 px for ${meas.length} elements incl. rotated, flipped and group-nested (worst ${worst.toFixed(2)} px)`, meas.filter((r) => r.worst > 1));
  ok(meas.some((r) => r.rot) && meas.some((r) => /flip/.test(r.name)) && meas.some((r) => /g0[678]/.test(r.name)), 'the checked set includes rotated, flipped and nested elements');
}
{
  await draw(SEL, 0);
  // ground truth by hand: rot-rect centre (860,100), 160x60, rotated 45 degrees; front-rect (200,140) 300x200, unrotated
  const g = await page.evaluate(() => {
    const { sel, sc, hnd } = globalThis.__s; const R = hnd.root.getBoundingClientRect(); const at = (n) => { sel.select(sc.elements.find((e) => e.name === n).id); return Object.fromEntries([...hnd.root.querySelectorAll('.sf-hdl')].map((h) => { const r = h.getBoundingClientRect(); return [h.dataset.h, { x: r.left + r.width / 2 - R.left, y: r.top + r.height / 2 - R.top, w: r.width, h: r.height }]; })); };
    const cs = (n) => getComputedStyle(hnd.root.querySelector('.sf-hdl'));
    const rot = at('rot-rect'); const front = at('front-rect'); const st = getComputedStyle(hnd.root.querySelector('.sf-hdl'));
    return { rot, front, style: { pe: st.pointerEvents, cursor: st.cursor, bg: st.backgroundColor, bw: st.borderTopWidth, w: st.width, shadow: st.boxShadow !== 'none' } };
  });
  const near = (a, x, y, tol = 0.75) => Math.hypot(a.x - x, a.y - y) <= tol;
  const k = Math.SQRT1_2; const c = [860, 100]; const along = (a, b) => [c[0] + (a * k - b * k), c[1] + (a * k + b * k)]; // (a,b) in the rect's own frame, rotated 45deg clockwise
  const want = { nw: along(-80, -30), ne: along(80, -30), se: along(80, 30), sw: along(-80, 30), n: along(0, -30), e: along(80, 0), s: along(0, 30), w: along(-80, 0) };
  ok(Object.entries(want).every(([h, [x, y]]) => near(g.rot[h], x, y)), 'rotated 45 degrees: all 8 handles are where the rotated rectangle\'s corners and edge midpoints are (hand-computed)', { got: g.rot, want });
  ok(g.rot.nw.w > 10 && g.rot.nw.w < 12.5, 'the handles themselves rotate with the element (an 8 px square at 45 degrees spans about 11.3 px)', g.rot.nw);
  const fw = { nw: [200, 140], n: [350, 140], ne: [500, 140], e: [500, 240], se: [500, 340], s: [350, 340], sw: [200, 340], w: [200, 240] };
  ok(Object.entries(fw).every(([h, [x, y]]) => near(g.front[h], x, y)) && g.front.nw.w === 8 && g.front.nw.h === 8, 'unrotated: corners and midpoints are on the frame, handles are 8 x 8 px', g.front);
  ok(g.style.pe === 'none' && g.style.cursor === 'auto' && g.style.w === '8px' && /255, 255, 255/.test(g.style.bg) && /^1(\.5)?px$/.test(g.style.bw) && g.style.shadow, 'handles are white 8 px squares with a 1.5 px accent border (the browser draws 1 px on 1x screens) and a soft shadow, no pointer events and no resize cursor (nothing is draggable yet)', g.style);
}
{
  // zoom: the same slide at other widths. Handle size stays 8 px; positions scale with the slide.
  await draw(SEL, 0);
  const z = await page.evaluate(async () => {
    const { sel, sc, hnd } = globalThis.__s; sel.select(sc.elements.find((e) => e.name === 'front-rect').id); const out = {};
    for (const w of [480, 1440, 960]) {
      hnd.setWidth(w); await new Promise((r) => requestAnimationFrame(() => requestAnimationFrame(r)));
      const R = hnd.root.getBoundingClientRect(); const u = w / 960; const get = (h) => { const r = hnd.root.querySelector(`.sf-hdl[data-h="${h}"]`).getBoundingClientRect(); return { x: r.left + r.width / 2 - R.left, y: r.top + r.height / 2 - R.top, w: r.width }; };
      const nw = get('nw'); const se = get('se'); out[w] = { dnw: Math.hypot(nw.x - 200 * u, nw.y - 140 * u), dse: Math.hypot(se.x - 500 * u, se.y - 340 * u), size: nw.w };
    }
    return out;
  });
  ok(Object.values(z).every((r) => r.dnw <= 0.75 && r.dse <= 0.75 && r.size === 8), 'at 50 %, 150 % and 100 % slide scale the handles stay on the frame corners and stay 8 px', z);
}
{
  // small frames: edge-centre handles are left out when they would pile onto the corner handles; all 8 exist in the DOM, 8 are visible on normal elements
  await draw(SEL, 0);
  const v = await page.evaluate(() => {
    const { sel, sc, hnd } = globalThis.__s; const vis = (n) => { sel.select(sc.elements.find((e) => e.name === n).id); const hs = [...hnd.root.querySelectorAll('.sf-hdl')]; return { dom: hs.length, shown: hs.filter((h) => getComputedStyle(h).display !== 'none').map((h) => h.dataset.h).sort().join() }; };
    return { front: vis('front-rect'), line: vis('h-line'), small: vis('g-child-a'), diag: vis('diag-line') };
  });
  ok(v.front.dom === 8 && v.front.shown === ['e', 'n', 'ne', 'nw', 's', 'se', 'sw', 'w'].join(), 'a normal element shows all 8 handles', v.front);
  ok(v.line.dom === 8 && v.line.shown === ['ne', 'nw', 'se', 'sw'].join(), 'a hairline (300 x 0 pt) shows its 4 corner handles only', v.line);
  ok(v.small.dom === 8 && v.small.shown.split(',').length === 8 && v.diag.shown.split(',').length === 8, 'a 60 pt square and a 200 x 120 diagonal frame still show all 8', v);
}
{
  // visual only: handles do not catch pointer events; picking is identical with and without handles on screen
  await draw(SEL, 0);
  const r = await page.evaluate(() => {
    const { sel, sc, hnd } = globalThis.__s; const R = hnd.root.getBoundingClientRect(); const front = sc.elements.find((e) => e.name === 'front-rect');
    const pts = []; for (let x = 150; x <= 560; x += 10) for (let y = 100; y <= 380; y += 10) pts.push([x, y]);
    const names = () => pts.map(([x, y]) => { const id = sel.pick(R.left + x, R.top + y); return id; });
    const before = names(); sel.select(front.id); const during = names();
    const hdl = [...hnd.root.querySelectorAll('.sf-hdl')].map((h) => { const b = h.getBoundingClientRect(); const top = document.elementsFromPoint(b.left + b.width / 2, b.top + b.height / 2); return top.some((n) => n.classList.contains('sf-hdl') || n.classList.contains('sf-sel') || n.classList.contains('sf-sel-h')); });
    return { same: JSON.stringify(before) === JSON.stringify(during), hit: hdl.some(Boolean), listeners: hnd.root.querySelectorAll('.sf-sel-h, .sf-hdl').length };
  });
  ok(r.same, 'picking over and around the selected element gives the same result with and without handles');
  ok(!r.hit, 'the browser never reports a handle (or the outline) under the pointer: they are transparent to input');
}

// ======================================================================================================
section('rotation handle (Phase 3.2.2-B): one visual knob 24 px above the top-centre handle, in the element\'s own frame');
{
  const exp = JSON.parse(readFileSync(path.join(FX, 'p31-geometry.expected.json'), 'utf8')).elements;
  await draw(path.join(FX, 'p31-geometry.pptx'), 0);
  const res = await page.evaluate((exp) => {
    const { sel, sc, hnd } = globalThis.__s; const out = []; const u = 960 / (sc.width / 12700); const R = hnd.root.getBoundingClientRect();
    const none = hnd.root.querySelectorAll('.sf-rot').length;
    const ctr = (el) => { const r = el.getBoundingClientRect(); return [r.left + r.width / 2 - R.left, r.top + r.height / 2 - R.top]; };
    for (const [name, e] of Object.entries(exp)) {
      if (e.kind === 'group') continue;
      const el = sc.elements.find((x) => x.name === name); sel.select(el.id);
      const knobs = hnd.root.querySelectorAll('.sf-sel .sf-rot'); const nh = ctr(hnd.root.querySelector('.sf-hdl[data-h="n"]')); const k = knobs[0] && ctr(knobs[0]);
      let cx = 0; let cy = 0; for (let i = 0; i < 8; i += 2) { cx += (e.corners[i] / 12700) * u / 4; cy += (e.corners[i + 1] / 12700) * u / 4; } // frame centre from the independent corners
      const dn = Math.hypot(nh[0] - cx, nh[1] - cy); const up = [(nh[0] - cx) / dn, (nh[1] - cy) / dn]; // outward direction through the top-centre handle
      const want = [nh[0] + up[0] * 24, nh[1] + up[1] * 24];
      const wpx = Math.hypot(e.corners[2] - e.corners[0], e.corners[3] - e.corners[1]) / 12700 * u; const hpx = Math.hypot(e.corners[4] - e.corners[2], e.corners[5] - e.corners[3]) / 12700 * u;
      out.push({ name, count: knobs.length, inside: !!knobs[0]?.parentElement?.classList.contains('sf-sel-h'), err: k ? Math.hypot(k[0] - want[0], k[1] - want[1]) : 99, gap: k ? Math.hypot(k[0] - nh[0], k[1] - nh[1]) : 0, big: wpx >= 40 && hpx >= 40, shear: !!el.world.shear, rot: el.world.rot });
    }
    sel.clear();
    return { out, none, after: hnd.root.querySelectorAll('.sf-rot').length };
  }, exp);
  ok(res.none === 0 && res.after === 0, 'no rotation handle exists before a selection and none is left after clear()');
  ok(res.out.every((r) => r.count === 1 && r.inside), `every selected element gets exactly one rotation handle, inside the handle layer (${res.out.length} elements)`, res.out.filter((r) => r.count !== 1));
  const meas = res.out.filter((r) => r.big && !r.shear);
  const worst = Math.max(...meas.map((r) => r.err)); const gaps = meas.map((r) => r.gap);
  ok(meas.length >= 8 && worst <= 1 && gaps.every((g) => Math.abs(g - 24) <= 0.75), `the knob is 24 px beyond the top-centre handle on the line from the frame centre through it, within 1 px, for ${meas.length} elements incl. rotated, flipped and nested (worst ${worst.toFixed(2)} px)`, meas.filter((r) => r.err > 1 || Math.abs(r.gap - 24) > 0.75));
  ok(meas.some((r) => r.rot) && meas.some((r) => /flip/.test(r.name)) && meas.some((r) => /g0[678]/.test(r.name)), 'the checked set includes rotated, flipped and nested elements');
}
{
  await draw(SEL, 0);
  const g = await page.evaluate(() => {
    const { sel, sc, hnd } = globalThis.__s; const R = hnd.root.getBoundingClientRect();
    const knob = (n) => { sel.select(sc.elements.find((e) => e.name === n).id); const el = hnd.root.querySelector('.sf-rot'); const r = el.getBoundingClientRect(); return { x: r.left + r.width / 2 - R.left, y: r.top + r.height / 2 - R.top, w: r.width }; };
    const front = knob('front-rect'); const rot = knob('rot-rect'); const line = knob('h-line');
    const k = hnd.root.querySelector('.sf-rot'); const h = hnd.root.querySelector('.sf-hdl'); const ks = getComputedStyle(k); const hs = getComputedStyle(h); const st = getComputedStyle(k, '::before');
    const same = ['width', 'height', 'backgroundColor', 'borderTopWidth', 'borderTopStyle', 'borderTopColor', 'boxShadow', 'pointerEvents', 'cursor'].every((p) => ks[p] === hs[p]);
    return { front, rot, line, same, round: ks.borderTopLeftRadius, stem: { w: st.width, h: st.height, bg: st.backgroundColor, accent: getComputedStyle(document.documentElement).getPropertyValue('--accent') }, pe: ks.pointerEvents, cursor: ks.cursor, stemPe: st.pointerEvents };
  });
  const k = Math.SQRT1_2; const c = [860, 100]; const along = (a, b) => [c[0] + (a * k - b * k), c[1] + (a * k + b * k)];
  const wantRot = along(0, -30 - 24); // 30 = half the 60 pt height, 24 px knob distance, in the rect's own frame, then rotated 45 degrees
  ok(Math.hypot(g.front.x - 350, g.front.y - (140 - 24)) <= 0.75, 'unrotated: the knob is centred above the top edge, 24 px up', g.front);
  ok(Math.hypot(g.rot.x - wantRot[0], g.rot.y - wantRot[1]) <= 0.75, 'rotated 45 degrees: the knob moves with the element, still 24 px out along its own "up" (hand-computed)', { got: g.rot, want: wantRot });
  ok(Math.abs(g.line.x - 210) <= 0.75 && Math.abs(g.line.y - (400 - 3 - 24)) <= 1, 'a zero-height line gets its knob above the (6 px minimum) outline', g.line);
  ok(g.same && g.round !== '0px', 'the knob matches the resize handles (size, fill, border, shadow, no input, no cursor); it is round to tell it apart', { same: g.same, round: g.round });
  ok(g.stem.w === '1px' && Math.abs(parseFloat(g.stem.h) - 20) <= 0.5 && g.stemPe === 'none', 'a 1 px accent stem joins the knob to the frame\'s top edge (20 px), also transparent to input', g.stem);
}
{
  // zoom: constant 24 px distance, position scales with the slide
  await draw(SEL, 0);
  const z = await page.evaluate(async () => {
    const { sel, sc, hnd } = globalThis.__s; sel.select(sc.elements.find((e) => e.name === 'front-rect').id); const out = {};
    for (const w of [480, 1440, 960]) {
      hnd.setWidth(w); await new Promise((r) => requestAnimationFrame(() => requestAnimationFrame(r)));
      const R = hnd.root.getBoundingClientRect(); const u = w / 960; const r = hnd.root.querySelector('.sf-rot').getBoundingClientRect();
      out[w] = { dx: r.left + r.width / 2 - R.left - 350 * u, dy: r.top + r.height / 2 - R.top - (140 * u - 24), size: r.width };
    }
    return out;
  });
  ok(Object.values(z).every((r) => Math.abs(r.dx) <= 0.75 && Math.abs(r.dy) <= 0.75 && r.size === 8), 'at 50 %, 150 % and 100 % slide scale the knob stays centred over the top edge and exactly 24 px away (it does not scale)', z);
}
{
  // not blocking input; present on small frames
  await draw(SEL, 0);
  const r = await page.evaluate(() => {
    const { sel, sc, hnd } = globalThis.__s; const R = hnd.root.getBoundingClientRect(); const front = sc.elements.find((e) => e.name === 'front-rect');
    const pts = []; for (let x = 150; x <= 560; x += 5) for (let y = 80; y <= 380; y += 5) pts.push([x, y]);
    const names = () => pts.map(([x, y]) => sel.pick(R.left + x, R.top + y));
    const before = names(); sel.select(front.id); const during = names();
    const b = hnd.root.querySelector('.sf-rot').getBoundingClientRect(); const top = document.elementsFromPoint(b.left + b.width / 2, b.top + b.height / 2);
    const stem = document.elementsFromPoint(b.left + b.width / 2, b.bottom + 6);
    const small = (n) => { sel.select(sc.elements.find((e) => e.name === n).id); const k = hnd.root.querySelector('.sf-rot'); return getComputedStyle(k).display !== 'none' && k.getBoundingClientRect().width > 0; };
    return { same: JSON.stringify(before) === JSON.stringify(during), hit: [...top, ...stem].some((n) => n.classList.contains('sf-rot') || n.classList.contains('sf-sel-h')), under: sc.byId.get(sel.pick(R.left + 350, R.top + 116))?.name, small: small('h-line') && small('g-child-a') };
  });
  ok(r.same, 'picking around the selected element is identical with and without the rotation handle on screen');
  ok(!r.hit, 'the browser never reports the knob or its stem under the pointer');
  ok(r.under === 'back-rect', 'the point under the knob belongs to the element behind it (back-rect), not to the knob', r);
  ok(r.small, 'the knob is shown for a hairline and a small element too');
}

// ======================================================================================================
section('nothing else changes: DOM identical until something is selected, identical again after clear, XML untouched (all fixtures and real decks)');
{
  const decks = [...['p31-geometry', 'p31-shapes', 'p31-text', 'p31-arabic', 'p31-media', 'p31-unsupported', 'p31-template-decor', 'p321-selection', 'p22-geometry'].map((n) => path.join(FX, n + '.pptx')), ...realFiles];
  for (const f of decks) {
    const r = await page.evaluate(async ({ b }) => {
      const { parsed, scene } = await globalThis.__load(b);
      const rn = await import('/js/render/render.js'); const sl = await import('/js/render/selection.js');
      const ser = (d) => new XMLSerializer().serializeToString(d);
      const out = { slides: 0, attachChanged: 0, pickChanged: 0, clearChanged: 0, overlayBad: 0, selected: 0, badPick: 0, picks: 0, mutated: 0, errors: [], strayProbe: 0, withText: 0 };
      for (let i = 0; i < parsed.slides.length; i++) {
        try {
          const sc = await scene.buildScene(parsed, i);
          const docs = [...new Set(sc.all.map((e) => e.el.ownerDocument))]; const before = docs.map(ser);
          const hnd = await rn.renderSlide(sc, document.getElementById('sf-host'), { width: 960 });
          const html0 = hnd.root.outerHTML;
          const sel = sl.attachSelection(hnd, sc, {});
          if (hnd.root.outerHTML !== html0) out.attachChanged++;
          // a grid of picks, every pointer type (exercises the hairline path as well)
          const R = hnd.root.getBoundingClientRect(); const okIds = new Set(sc.elements.filter((e) => e.kind !== 'group' && e.kind !== 'alt' && hnd.nodes.has(e.id)).map((e) => e.id));
          for (const t of ['mouse', 'touch']) for (let gx = 1; gx < 24; gx++) for (let gy = 1; gy < 14; gy++) {
            const id = sel.pick(R.left + (gx * R.width) / 24, R.top + (gy * R.height) / 14, t); out.picks++;
            if (id && !okIds.has(id)) out.badPick++;
          }
          if (hnd.root.outerHTML !== html0) out.pickChanged++;
          if (hnd.root.querySelector('svg:not(.sf-geo)')) out.strayProbe++;
          const first = sc.elements.find((e) => okIds.has(e.id));
          if (first) {
            sel.select(first.id); out.selected++;
            if (hnd.root.querySelectorAll('.sf-sel').length !== 1 || hnd.root.lastElementChild.className !== 'sf-sel') out.overlayBad++;
            sel.clear();
            if (hnd.root.outerHTML !== html0) out.clearChanged++;
          }
          docs.forEach((d, j) => { if (ser(d) !== before[j]) out.mutated++; });
          sel.dispose(); out.slides++;
        } catch (e) { out.errors.push(i + ': ' + String(e).slice(0, 200)); }
      }
      return out;
    }, { b: b64(f) });
    const nm = path.basename(f);
    ok(r.errors.length === 0 && r.slides > 0, `${nm}: ${r.slides} slide(s) take selection and ${r.picks} grid picks without errors`, r.errors);
    ok(r.attachChanged === 0 && r.pickChanged === 0 && r.clearChanged === 0 && r.strayProbe === 0, `${nm}: the drawn DOM is byte-identical after attaching, after picking, and after select + clear`, r);
    ok(r.overlayBad === 0 && r.badPick === 0, `${nm}: one outline on top of the slide while selected; every pick is a selectable slide element`, r);
    ok(r.mutated === 0, `${nm}: slide / layout / master XML is identical after selection (read-only)`, r);
  }
}

// ======================================================================================================
section('code guards: selection is read-only, CSP-safe and independent of storage');
{
  const code = readFileSync(path.join(ROOT, 'js/render/selection.js'), 'utf8').replace(/\/\/.*$/gm, '');
  ok(!/\bimport\b[^;]*(store|firebase|firestore|auth)/i.test(code) && !/indexedDB|localStorage|sessionStorage|\bfetch\(|XMLHttpRequest|XMLSerializer|navigator\.sendBeacon/.test(code), 'selection.js imports no store / Firebase and has no storage, network or XML serialisation');
  ok(!/innerHTML|outerHTML|insertAdjacentHTML|\beval\(|new Function|setAttribute\(\s*['"]style|cssText/.test(code), 'selection.js uses no innerHTML, eval, inline style attribute or cssText (CSP: style-src \'self\')');
  ok(!/\.el\b/.test(code) && !/\.doc\b|ownerDocument\.(?!elementsFromPoint)/.test(code), 'selection.js never reaches the XML nodes (no scene element .el or .doc access)');
  const css = readFileSync(path.join(ROOT, 'css/render.css'), 'utf8');
  ok(/\.sf-sel\s*\{[^}]*pointer-events:\s*none/.test(css), 'the outline rule is static CSS with pointer-events: none');
}

// ======================================================================================================
// the real UI: upload, preview, Rendered tab, real mouse and touch input
section('Rendered tab: mouse, in the real UI');
async function openRendered(pg, name) {
  await pg.goto(BASE + '/index.html'); await pg.waitForSelector('.shell');
  await pg.click('text=Templates >> nth=0'); await pg.waitForSelector('.dropzone');
  const [fc] = await Promise.all([pg.waitForEvent('filechooser'), pg.click('.dropzone')]); await fc.setFiles(SEL);
  await pg.waitForSelector('.modal #tpl-name'); await pg.fill('#tpl-name', name); await pg.click('.modal >> text=Save to library');
  await pg.waitForFunction((n) => [...document.querySelectorAll('.tpl h3')].some((x) => x.textContent === n), name);
  await pg.click(`.tpl:has-text("${name}") >> text=Preview`); await pg.waitForSelector('.film-item');
  await pg.click('role=tab[name="Rendered"]'); await pg.waitForSelector('.rendered-host .sf-slide .sf-el');
}
const slideBox = (pg) => pg.evaluate(() => { const r = document.querySelector('.rendered-host .sf-slide').getBoundingClientRect(); return { x: r.left, y: r.top, s: r.width / 960 }; });
const at = async (pg, x, y) => { const b = await slideBox(pg); return [b.x + x * b.s, b.y + y * b.s]; };
const state = (pg) => pg.evaluate(() => {
  const root = document.querySelector('.rendered-host .sf-slide'); const o = [...root.querySelectorAll('.sf-sel')];
  const node = o[0] && [...root.querySelectorAll('.sf-el')].find((n) => n.dataset.id === o[0].dataset.for);
  const a = o[0]?.getBoundingClientRect(); const b = node?.getBoundingClientRect();
  return { count: o.length, id: o[0]?.dataset.for || null, name: node?.dataset.name || null, status: document.querySelector('.sf-sel-status')?.textContent || '', slide: root.dataset.slide, dx: a && b ? Math.max(Math.abs(a.left - b.left), Math.abs(a.right - b.right), Math.abs(a.top - b.top), Math.abs(a.bottom - b.bottom)) : null };
});
{
  const ui = await ctx.newPage(); watch(ui, 'ui: ');
  await ui.setViewportSize({ width: 1280, height: 900 });
  await openRendered(ui, 'Selection deck');
  let s = await state(ui);
  ok(s.count === 0 && /Nothing selected/.test(s.status), 'nothing is selected when the Rendered tab opens, and the status line says how to select', s);
  ok(await ui.evaluate(() => !document.querySelector('.rendered-host input, .rendered-host textarea, .rendered-host [contenteditable], .rendered-host button[type=submit]') && !document.querySelector('.sf-sel-status button')), 'still read-only: no inputs, no toolbar, no edit or save controls were added');

  await ui.mouse.click(...(await at(ui, 450, 300)));
  s = await state(ui);
  ok(s.count === 1 && s.name === 'front-rect' && s.id === 'ppt/slides/slide1.xml#3', 'a click selects the element under the pointer by its Scene Model id', s);
  ok(/Selected: Shape “front-rect” \(ppt\/slides\/slide1\.xml#3\)/.test(s.status), 'the status line names the selected element and its id', s.status);
  ok(s.dx !== null && s.dx <= 1, 'the outline matches the element\'s rectangle at the scale on screen (1 px)', s);
  const hd = () => ui.evaluate(() => { const root = document.querySelector('.rendered-host .sf-slide'); const R = root.getBoundingClientRect(); const o = root.querySelector('.sf-sel'); const f = o && [...root.querySelectorAll('.sf-el')].find((n) => n.dataset.id === o.dataset.for)?.getBoundingClientRect(); const get = (h) => { const r = root.querySelector(`.sf-hdl[data-h="${h}"]`)?.getBoundingClientRect(); return r && { x: r.left + r.width / 2, y: r.top + r.height / 2 }; }; return { n: root.querySelectorAll('.sf-hdl').length, nw: get('nw'), se: get('se'), f: f && { l: f.left, t: f.top, r: f.right, b: f.bottom } }; });
  let hh = await hd();
  ok(hh.n === 8 && Math.hypot(hh.nw.x - hh.f.l, hh.nw.y - hh.f.t) <= 1 && Math.hypot(hh.se.x - hh.f.r, hh.se.y - hh.f.b) <= 1, 'selecting in the real UI shows 8 handles, the corner ones on the element\'s corners at the scale on screen', hh);
  await ui.mouse.click(...(await at(ui, 196, 240)));
  s = await state(ui);
  ok(s.name === 'back-rect', 'a click on the outer half of a handle goes through to the element under it (handles never intercept clicks)', s);
  await ui.mouse.click(...(await at(ui, 450, 300)));
  const rotAt = () => ui.evaluate(() => { const root = document.querySelector('.rendered-host .sf-slide'); const k = root.querySelector('.sf-rot'); const n = root.querySelector('.sf-hdl[data-h="n"]'); const c = (e) => { const r = e.getBoundingClientRect(); return { x: r.left + r.width / 2, y: r.top + r.height / 2 }; }; return { count: root.querySelectorAll('.sf-rot').length, k: k && c(k), n: n && c(n) }; });
  let rk = await rotAt(); const bx = await slideBox(ui);
  ok(rk.count === 1 && Math.abs(rk.k.x - rk.n.x) <= 0.75 && Math.abs(rk.n.y - rk.k.y - 24) <= 0.75, 'in the real UI the rotation handle is 24 px above the top-centre handle', rk);
  await ui.mouse.click(...(await at(ui, 350, 116)));
  s = await state(ui); ok(s.name === 'back-rect', 'a click on the rotation handle goes through to the element behind it', s);
  await ui.mouse.click(...(await at(ui, 450, 300)));

  await ui.mouse.click(...(await at(ui, 100, 100)));
  s = await state(ui);
  ok(s.count === 1 && s.name === 'back-rect', 'clicking another element updates the selection (still exactly one outline)', s);
  await ui.mouse.click(...(await at(ui, 250, 200)));
  ok((await state(ui)).name === 'front-rect', 'overlapping elements: the visually higher one is selected');
  await ui.mouse.click(...(await at(ui, 530, 70)));
  ok((await state(ui)).name === 'under-rect', 'a click in the empty corner of a circle\'s frame selects what is visible there');
  await ui.mouse.click(...(await at(ui, 860, 100)));
  s = await state(ui);
  ok(s.name === 'rot-rect' && s.dx <= 1.5, 'a rotated element is selected and outlined along its rotation', s);

  await ui.mouse.click(...(await at(ui, 30, 30)));
  s = await state(ui);
  ok(s.count === 0 && /Nothing selected/.test(s.status), 'clicking empty space clears the selection and its outline');
  ok((await hd()).n === 0 && (await rotAt()).count === 0, 'and the handles and the rotation handle are gone with it');
  await ui.mouse.click(...(await at(ui, 450, 300))); await ui.mouse.click(...(await at(ui, 500, 525)));
  ok((await state(ui)).count === 0, 'clicking master decoration counts as empty space');
  await ui.mouse.click(...(await at(ui, 110, 500)));
  ok((await state(ui)).count === 0, 'clicking where a hidden element is does nothing');

  // hairline through the real mouse
  const [hx, hy] = await at(ui, 200, 402); await ui.mouse.click(hx, hy);
  s = await state(ui);
  ok(s.name === 'h-line', 'a 1 pt line can be selected with the mouse 2 px off its centre', s);

  // a press that travels is a drag (e.g. selecting text), not a click
  await ui.mouse.click(...(await at(ui, 450, 300)));
  const [dx0, dy0] = await at(ui, 100, 100);
  await ui.mouse.move(dx0, dy0); await ui.mouse.down(); await ui.mouse.move(dx0 + 40, dy0 + 30, { steps: 5 }); await ui.mouse.up();
  ok((await state(ui)).name === 'front-rect', 'dragging across the slide does not change the selection');
  await ui.mouse.click(...(await at(ui, 100, 100)), { button: 'right' });
  ok((await state(ui)).name === 'front-rect', 'a right click does not change the selection');

  // text can still be selected and copied
  await ui.mouse.dblclick(...(await at(ui, 90, 320)));
  const txt = await ui.evaluate(() => getSelection().toString());
  ok(/Label/.test(txt), 'text in the slide can still be selected with the mouse (selection did not take over text selection)', txt);
  ok((await state(ui)).name === 'textbox', 'and the element holding that text is the selected element');

  // switching slides
  await ui.mouse.click(...(await at(ui, 450, 300)));
  await ui.click('.rendered .film-item >> nth=1'); await ui.waitForSelector('.rendered-host .sf-slide[data-slide="1"]');
  s = await state(ui);
  ok(s.count === 0 && /Nothing selected/.test(s.status), 'switching slide clears the selection (filmstrip)', s);
  ok((await hd()).n === 0 && (await rotAt()).count === 0, 'and no handles or rotation handle are left on the new slide');
  await ui.mouse.click(...(await at(ui, 150, 150)));
  s = await state(ui);
  ok(s.name === 's2-rect' && s.id === 'ppt/slides/slide2.xml#2', 'on slide 2 the same cNvPr id number is a different Scene Model id (qualified by the slide part)', s);
  await ui.keyboard.press('ArrowUp'); await ui.waitForSelector('.rendered-host .sf-slide[data-slide="0"]');
  s = await state(ui);
  ok(s.count === 0, 'switching slide with the keyboard also clears it, and slide 1 does not inherit slide 2\'s selection', s);
  await ui.keyboard.press('ArrowDown'); await ui.waitForSelector('.rendered-host .sf-slide[data-slide="1"]');
  ok((await state(ui)).count === 0, 'coming back to slide 2 starts without a selection');
  await ui.click('.rendered .film-item >> nth=0'); await ui.waitForSelector('.rendered-host .sf-slide[data-slide="0"]');

  // other tabs and back: the same slide is drawn again, its selection is kept
  await ui.mouse.click(...(await at(ui, 450, 300)));
  await ui.click('role=tab[name="Slides"]'); await ui.click('role=tab[name="Rendered"]'); await ui.waitForSelector('.rendered-host .sf-slide[data-slide="0"] .sf-el');
  s = await state(ui);
  ok(s.name === 'front-rect' && s.count === 1, 'leaving the Rendered tab and coming back keeps the selection of the same slide', s);
  await ui.click('role=tab[name="Slides"]'); await ui.click('.preview:not(.rendered) .film-item >> nth=1');
  await ui.click('role=tab[name="Rendered"]'); await ui.waitForSelector('.rendered-host .sf-slide[data-slide="1"] .sf-el');
  ok((await state(ui)).count === 0, 'choosing another slide in the Slides tab clears it as well');
  await ui.click('.rendered .film-item >> nth=0'); await ui.waitForSelector('.rendered-host .sf-slide[data-slide="0"]');
  ok((await state(ui)).count === 0, 'and returning to slide 1 does not bring the old selection back');

  // resize: the outline follows (it is positioned with the same --u variable as the slide)
  await ui.mouse.click(...(await at(ui, 860, 100)));
  await ui.setViewportSize({ width: 820, height: 900 });
  await ui.waitForFunction(() => { const r = document.querySelector('.rendered-host .sf-slide'); const f = document.querySelector('.sf-frame'); return r && Math.abs(r.getBoundingClientRect().width - f.clientWidth) <= 1 && f.clientWidth < 800; });
  s = await state(ui);
  ok(s.name === 'rot-rect' && s.dx <= 1.5, 'after the window is resized the outline still matches the element', s);
  hh = await hd();
  const b2 = await slideBox(ui); const k2 = Math.SQRT1_2;
  rk = await rotAt();
  { const s2 = b2.s; const px = b2.x + 860 * s2; const py = b2.y + 100 * s2; const up = 30 * s2 + 24; const wx = px + up * k2; const wy = py - up * k2;
    ok(rk.count === 1 && Math.hypot(rk.k.x - wx, rk.k.y - wy) <= 1.25, 'and the rotation handle of the rotated element is still 24 px out from its top edge at the new scale (hand-computed)', { rk, want: [wx, wy] }); }
  ok(hh.n === 8 && hh.nw && Math.hypot(hh.nw.x - (b2.x + (860 + (-80 * k2 + 30 * k2)) * b2.s), hh.nw.y - (b2.y + (100 + (-80 * k2 - 30 * k2)) * b2.s)) <= 1.25, 'and the handles of a rotated element follow it at the new scale (nw handle vs hand-computed corner)', { hh, b2 });
  await ui.mouse.click(...(await at(ui, 450, 300)));
  ok((await state(ui)).name === 'front-rect', 'and hit-testing still lands on the right element at the new scale');
  await ui.setViewportSize({ width: 1280, height: 900 });

  ok(await ui.evaluate(() => document.querySelectorAll('canvas').length === 0), 'still no canvas / raster in the Rendered tab');
  await ui.keyboard.press('Escape');
  await ui.close();
}

// ======================================================================================================
section('Rendered tab: touch (phone width, real touch input)');
{
  const tctx = await browser.newContext({ viewport: { width: 1280, height: 900 }, hasTouch: true, deviceScaleFactor: 2 });
  const ui = await tctx.newPage(); watch(ui, 'touch: ');
  await openRendered(ui, 'Selection deck');
  await ui.setViewportSize({ width: 390, height: 844 });
  await ui.waitForFunction(() => document.querySelector('.rendered-host .sf-slide').getBoundingClientRect().width < 380);
  const tap = async (x, y) => { const [px, py] = await at(ui, x, y); await ui.touchscreen.tap(px, py); };
  const b = await slideBox(ui);
  ok(b.s < 0.5, 'the slide is scaled down for the phone width', b);
  await ui.evaluate(() => document.querySelector('.rendered-host .sf-slide').scrollIntoView({ block: 'center' }));
  await tap(450, 300);
  let s = await state(ui);
  ok(s.name === 'front-rect' && s.count === 1 && s.dx <= 1, 'a tap selects the element and draws its outline', s);
  const th = await ui.evaluate(() => { const root = document.querySelector('.rendered-host .sf-slide'); const h = root.querySelector('.sf-hdl'); return { n: root.querySelectorAll('.sf-hdl').length, size: h.getBoundingClientRect().width, pe: getComputedStyle(h).pointerEvents }; });
  ok(th.n === 8 && th.pe === 'none' && th.size >= 8, 'a tap shows 8 handles at phone width, still transparent to input', th);
  const tk = await ui.evaluate(() => { const k = document.querySelector('.rendered-host .sf-slide .sf-rot'); return { n: document.querySelectorAll('.rendered-host .sf-rot').length, pe: getComputedStyle(k).pointerEvents }; });
  ok(tk.n === 1 && tk.pe === 'none', 'a tap also shows the rotation handle, transparent to input', tk);
  await tap(200, 240);
  s = await state(ui);
  ok(s.name === 'front-rect' || s.name === 'back-rect', 'tapping on a handle location behaves like tapping the slide there (no handle interaction)', s);
  await tap(450, 300);
  await tap(100, 100);
  ok((await state(ui)).name === 'back-rect', 'tapping another element updates the selection');
  await tap(250, 200);
  ok((await state(ui)).name === 'front-rect', 'overlapping elements: tap selects the visually higher one');
  await tap(30, 30);
  ok((await state(ui)).count === 0, 'tapping empty space clears the selection');
  await tap(200, 407);
  ok((await state(ui)).name === 'h-line', 'a finger 7 pt beside a hairline still selects the line (touch reach)');
  await tap(530, 70);
  ok((await state(ui)).name === 'under-rect', 'tap on the empty corner of the circle\'s frame selects what is visible');

  // a touch that moves (a scroll or a swipe) is not a tap
  await tap(450, 300);
  const cdp = await tctx.newCDPSession(ui);
  const [sx, sy] = await at(ui, 100, 100);
  await cdp.send('Input.dispatchTouchEvent', { type: 'touchStart', touchPoints: [{ x: sx, y: sy }] });
  await cdp.send('Input.dispatchTouchEvent', { type: 'touchMove', touchPoints: [{ x: sx, y: sy - 60 }] });
  await cdp.send('Input.dispatchTouchEvent', { type: 'touchEnd', touchPoints: [] });
  await ui.waitForTimeout(150);
  ok((await state(ui)).name === 'front-rect', 'a swipe that starts on another element does not select it');
  // two fingers (pinch) are not a tap either
  const [ax, ay] = await at(ui, 100, 100); const [bx, by] = await at(ui, 250, 120);
  await cdp.send('Input.dispatchTouchEvent', { type: 'touchStart', touchPoints: [{ x: ax, y: ay, id: 1 }, { x: bx, y: by, id: 2 }] });
  await cdp.send('Input.dispatchTouchEvent', { type: 'touchEnd', touchPoints: [] });
  await ui.waitForTimeout(150);
  ok((await state(ui)).name === 'front-rect', 'a two-finger touch does not change the selection');
  ok(await ui.evaluate(() => getComputedStyle(document.querySelector('.rendered-host .sf-slide')).touchAction === 'auto'), 'the slide keeps touch-action: auto (scroll and pinch-zoom still work)');
  await tctx.close();
}

// ======================================================================================================
section('move preview (Phase 3.2.3-A): drag the selected element, temporary, preview only');
await page.evaluate(() => {
  // synthetic pointer events on the slide root, in slide px (the harness slide is 960 px wide at the top-left: 1 pt = 1 px). Mouse presses carry buttons: 1 until released.
  globalThis.__ptr = (type, x, y, o = {}) => {
    const root = globalThis.__s.hnd.root; const R = root.getBoundingClientRect();
    return root.dispatchEvent(new PointerEvent(type, { bubbles: true, cancelable: true, composed: true, pointerType: 'mouse', isPrimary: true, pointerId: 1, button: 0, buttons: type === 'pointerup' ? 0 : 1, clientX: R.left + x, clientY: R.top + y, ...o }));
  };
  // everything a test needs to know about the selected element right now (positions relative to the slide, px)
  globalThis.__snap = () => {
    const { sel, hnd } = globalThis.__s; const id = sel.id; if (!id) return null;
    const R = hnd.root.getBoundingClientRect(); const rc = (n) => { const r = n.getBoundingClientRect(); return { l: r.left - R.left, t: r.top - R.top, r: r.right - R.left, b: r.bottom - R.top }; };
    const node = hnd.nodes.get(id); const o = hnd.root.querySelector('.sf-sel');
    const cen = (q) => [...hnd.root.querySelectorAll(q)].map((h) => { const r = h.getBoundingClientRect(); return [r.left + r.width / 2 - R.left, r.top + r.height / 2 - R.top]; });
    const geo = node.querySelector('.sf-geo');
    return { id, node: rc(node), out: o && rc(o), hs: cen('.sf-sel .sf-hdl'), knob: cen('.sf-sel .sf-rot'), tf: getComputedStyle(node).transform, geoTf: geo ? getComputedStyle(geo).transform : '', offset: sel.offset(id), dragging: sel.dragging, selId: sel.id };
  };
});
const snap = () => page.evaluate(() => globalThis.__snap());
const shift = (a, b) => ({ dl: b.l - a.l, dt: b.t - a.t, dr: b.r - a.r, db: b.b - a.b });
const near = (v, w, tol = 0.6) => Math.abs(v - w) <= tol;
const selectByName = (n) => page.evaluate((n) => { const { sel, sc } = globalThis.__s; return sel.select(sc.elements.find((e) => e.name === n).id); }, n);
const worldKey = () => page.evaluate(() => { const { sc } = globalThis.__s; const ser = (d) => new XMLSerializer().serializeToString(d); const docs = [...new Set(sc.all.map((e) => e.el.ownerDocument))]; return JSON.stringify({ w: sc.all.map((e) => [e.id, e.world]), x: docs.map(ser) }); });

// ---- 1. a real mouse drag on the selected element -------------------------------------------------
{
  await draw(SEL, 0);
  const html0 = await page.evaluate(() => globalThis.__s.hnd.root.outerHTML);
  const key0 = await worldKey();
  await selectByName('front-rect');
  const a = await snap();
  ok(a.offset.dx === 0 && a.offset.dy === 0 && a.dragging === null, 'a selected element starts with no preview offset and is not being dragged', a.offset);
  await page.mouse.move(350, 240); await page.mouse.down(); await page.mouse.move(410, 280, { steps: 6 });
  const mid = await snap();
  ok(mid.dragging === a.id && near(shift(a.node, mid.node).dl, 60) && near(shift(a.node, mid.node).dt, 40), 'during the drag the element follows the pointer exactly (60 px right, 40 px down)', { mid: mid.node, from: a.node });
  await page.mouse.up();
  const b = await snap();
  const d = shift(a.node, b.node);
  ok(b.dragging === null && near(d.dl, 60) && near(d.dr, 60) && near(d.dt, 40) && near(d.db, 40), 'after the release the element stays 60 px right and 40 px down: same size, only moved', d);
  ok(near(b.offset.dx, 60) && near(b.offset.dy, 40), 'the preview offset is reported in slide points (1 px = 1 pt here)', b.offset);
  const od = shift(b.node, b.out);
  ok(Math.max(...Object.values(od).map(Math.abs)) <= 1, 'the selection outline moved with the element and still matches its frame (1 px)', od);
  ok(b.hs.every((p, i) => near(p[0] - a.hs[i][0], 60) && near(p[1] - a.hs[i][1], 40)) && b.knob.every((p, i) => near(p[0] - a.knob[i][0], 60) && near(p[1] - a.knob[i][1], 40)), 'all 8 resize handles and the rotation handle moved by exactly the same 60 / 40 px', { hs: b.hs[0], from: a.hs[0] });
  ok(b.selId === a.id, 'the element is still the selected one (a drag is not a click and does not reselect)');
  const st = await page.evaluate(() => { const { hnd, changes, moves } = globalThis.__s; return { changes: changes.length, moves: moves.slice(), cur: getComputedStyle(hnd.nodes.get(globalThis.__s.sel.id)).touchAction, us: hnd.root.style.userSelect, sheet: !!hnd.root.querySelector('.sf-sel') }; });
  ok(st.changes === 1 && JSON.stringify(st.moves) === '["start","end"]', 'onChange is not fired by the drag; onMove reports start and end once', st);
  ok(st.us === '', 'text selection is switched off only while dragging', st);
  ok((await worldKey()) === key0, 'Scene Model positions (every element\'s world frame) and the slide / layout / master XML are identical after the move: preview only');
  // hit-testing follows what is drawn now
  const pk = await page.evaluate(() => { const { sel, sc } = globalThis.__s; const n = (x, y) => { const id = sel.pick(x, y); return id ? sc.byId.get(id).name : null; }; return { newCentre: n(410, 280), farEdge: n(550, 300), oldOnly: n(230, 300), oldLeftTop: n(210, 150) }; });
  ok(pk.newCentre === 'front-rect' && pk.farEdge === 'front-rect' && pk.oldOnly !== 'front-rect' && pk.oldLeftTop !== 'front-rect', 'picking follows the moved element: it is found at its new place and no longer at the old one', pk);
  await page.evaluate(() => { globalThis.__s.sel.reset(); });
  const r = await snap();
  ok(r.offset.dx === 0 && near(shift(a.node, r.node).dl, 0, 0.01) && near(shift(a.node, r.node).dt, 0, 0.01) && r.out && near(shift(a.out, r.out).dl, 0, 0.01), 'reset() puts the element and its outline back exactly');
  await page.evaluate(() => globalThis.__s.sel.clear());
  ok(await page.evaluate((h) => globalThis.__s.hnd.root.outerHTML === h, html0), 'after reset + clear the drawn DOM is byte-identical to the unselected slide (no touch-action, cursor, user-select or translate left behind)');
}

// ---- 2. only the selected element moves; small movements are still clicks ---------------------------
{
  await draw(SEL, 0);
  await selectByName('front-rect');
  const any = () => page.evaluate(() => [...globalThis.__s.hnd.root.querySelectorAll('.sf-el')].filter((n) => n.style.translate).length);
  await page.mouse.move(100, 100); await page.mouse.down(); await page.mouse.move(160, 150, { steps: 5 }); await page.mouse.up();
  ok((await any()) === 0 && (await snap()).selId === (await snap()).id, 'pressing on another (unselected) element and dragging moves nothing and does not change the selection');
  await page.mouse.move(30, 480); await page.mouse.down(); await page.mouse.move(80, 500, { steps: 5 }); await page.mouse.up();
  ok((await any()) === 0, 'dragging from empty space moves nothing');
  await page.mouse.move(350, 240); await page.mouse.down(); await page.mouse.move(353, 240); await page.mouse.up();
  ok((await any()) === 0 && (await snap()).offset.dx === 0, 'a press on the selected element that moves 3 px (mouse click limit is 4) is a click: no move');
  await page.mouse.move(350, 240); await page.mouse.down(); await page.mouse.move(356, 240); await page.mouse.up();
  ok(near((await snap()).offset.dx, 6), 'moving 6 px starts the drag and the element goes the full 6 px (no lag at the start)', (await snap()).offset);
  // the same press-and-release on the element that is under another one: the visually higher one decides
  await draw(SEL, 0); await selectByName('back-rect');
  await page.mouse.move(250, 200); await page.mouse.down(); await page.mouse.move(300, 230, { steps: 4 }); await page.mouse.up();
  ok((await any()) === 0, 'pressing where a higher element covers the selected one does not drag the selected one (the covering element is what is under the pointer)');
}

// ---- 3. zoom: same screen delta, any scale; the offset is stored in points ---------------------------
{
  await draw(SEL, 0);
  await selectByName('front-rect');
  const rz = () => page.evaluate(() => new Promise((r) => requestAnimationFrame(() => requestAnimationFrame(r))));
  const res = {};
  for (const w of [480, 960, 1440]) {
    await draw(SEL, 0); await selectByName('front-rect');
    await page.evaluate((w) => globalThis.__s.hnd.setWidth(w), w); await rz();
    const u = w / 960; const a = await snap();
    await page.mouse.move(350 * u, 240 * u); await page.mouse.down(); await page.mouse.move(350 * u + 60, 240 * u + 40, { steps: 5 }); await page.mouse.up();
    const b = await snap();
    res[w] = { dx: shift(a.node, b.node).dl, dy: shift(a.node, b.node).dt, off: b.offset, wantPt: [60 / u, 40 / u], outDx: shift(b.node, b.out) };
  }
  ok(Object.entries(res).every(([w, r]) => near(r.dx, 60) && near(r.dy, 40) && near(r.off.dx, r.wantPt[0], 0.05) && near(r.off.dy, r.wantPt[1], 0.05)), 'a 60 / 40 px drag moves the element 60 / 40 px at 50 %, 100 % and 150 % slide scale, and the offset is 120 / 80, 60 / 40 and 40 / 26.7 slide points', res);
  ok(Object.values(res).every((r) => Math.max(...Object.values(r.outDx).map(Math.abs)) <= 1), 'the outline stays on the moved element at every scale (1 px)', res);
  // the offset is in points: resizing the slide afterwards scales the move with it
  await draw(SEL, 0); await selectByName('front-rect');
  await page.mouse.move(350, 240); await page.mouse.down(); await page.mouse.move(410, 280, { steps: 5 }); await page.mouse.up();
  const base = await snap();
  const after = {};
  for (const w of [480, 1440, 960]) { await page.evaluate((w) => globalThis.__s.hnd.setWidth(w), w); await rz(); const s = await snap(); const u = w / 960; after[w] = { l: s.node.l - 200 * u, t: s.node.t - 140 * u, out: Math.max(...Object.values(shift(s.node, s.out)).map(Math.abs)) }; }
  ok(Object.entries(after).every(([w, r]) => near(r.l, 60 * (w / 960), 0.8) && near(r.t, 40 * (w / 960), 0.8) && r.out <= 1), 'resizing the slide after the move keeps the element and its outline at the same place on the slide (50 %, 150 %, 100 %)', after);
  // zoom changes WHILE dragging: the element still goes where the pointer is, measured on the slide as drawn now
  await draw(SEL, 0); await selectByName('front-rect');
  await page.mouse.move(350, 240); await page.mouse.down(); await page.mouse.move(380, 240, { steps: 3 });
  await page.evaluate(() => globalThis.__s.hnd.setWidth(480)); await rz();
  await page.mouse.move((350 + 100) * 0.5, 240 * 0.5, { steps: 3 }); await page.mouse.up();
  const zd = await snap();
  ok(near(zd.offset.dx, 100, 0.1) && near(zd.offset.dy, 0, 0.1), 'the slide is zoomed out while the element is held: it ends exactly 100 slide points right of where it started', zd.offset);
}

// ---- 4. rotation, flips and group transforms, against the independently computed corners ------------
{
  const exp = JSON.parse(readFileSync(path.join(FX, 'p31-geometry.expected.json'), 'utf8')).elements;
  await draw(path.join(FX, 'p31-geometry.pptx'), 0);
  const names = Object.entries(exp).filter(([, e]) => e.kind !== 'group').map(([n]) => n);
  const rows = []; const skipped = [];
  for (const name of names) {
    // a point where the element itself is what is under the pointer (not every frame is painted at its centre)
    const spot = await page.evaluate((name) => {
      const { sel, sc, hnd } = globalThis.__s; const e = sc.elements.find((x) => x.name === name); sel.select(e.id);
      const R = hnd.root.getBoundingClientRect(); const r = hnd.nodes.get(e.id).getBoundingClientRect(); let best = null;
      for (let i = 1; i < 20; i++) for (let j = 1; j < 20; j++) { const x = r.left + (r.width * i) / 20; const y = r.top + (r.height * j) / 20; if (sel.pick(x, y) === e.id) { const d = Math.hypot(x - (r.left + r.width / 2), y - (r.top + r.height / 2)); if (!best || d < best.d) best = { x: x - R.left, y: y - R.top, d }; } }
      return best;
    }, name);
    if (!spot) { skipped.push(name); continue; }
    const a = await snap();
    await page.mouse.move(spot.x, spot.y); await page.mouse.down(); await page.mouse.move(spot.x + 37, spot.y + 23, { steps: 5 }); await page.mouse.up();
    const b = await snap();
    const e = exp[name]; const u = 960 / ((await page.evaluate(() => globalThis.__s.sc.width)) / 12700);
    const c = []; for (let i = 0; i < 8; i += 2) c.push([(e.corners[i] / 12700) * u + 37, (e.corners[i + 1] / 12700) * u + 23]);
    const mid = (p, q) => [(p[0] + q[0]) / 2, (p[1] + q[1]) / 2];
    const want = [c[0], c[1], c[2], c[3], mid(c[0], c[1]), mid(c[1], c[2]), mid(c[2], c[3]), mid(c[3], c[0])];
    const used = new Set(); let worst = 0;
    for (const w of want) { let bi = -1; let bd = 1e9; b.hs.forEach((p, i) => { const dd = Math.hypot(p[0] - w[0], p[1] - w[1]); if (!used.has(i) && dd < bd) { bd = dd; bi = i; } }); used.add(bi); worst = Math.max(worst, bd); }
    const wpx = Math.hypot(c[1][0] - c[0][0], c[1][1] - c[0][1]); const hpx = Math.hypot(c[2][0] - c[1][0], c[2][1] - c[1][1]);
    const os = shift(a.out, b.out); rows.push({ name, dl: shift(a.node, b.node).dl, dt: shift(a.node, b.node).dt, dOut: Math.max(Math.abs(os.dl - 37), Math.abs(os.dr - 37), Math.abs(os.dt - 23), Math.abs(os.db - 23)), worst, sameTf: a.tf === b.tf && a.geoTf === b.geoTf, big: wpx >= 40 && hpx >= 40, kind: e.kind, hs: b.hs.length });
  }
  ok(rows.length >= 12, `${rows.length} elements of p31-geometry (plain, rotated, flipped, group-nested, lines, text) could be grabbed and dragged with a real mouse`, { skipped });
  ok(rows.every((r) => near(r.dl, 37, 0.8) && near(r.dt, 23, 0.8)), 'each moved exactly 37 px right and 23 px down in slide space, whatever its rotation, flip or group', rows.filter((r) => !(near(r.dl, 37, 0.8) && near(r.dt, 23, 0.8))));
  ok(rows.every((r) => r.sameTf), 'the rotation and flip transforms of every dragged element (frame and drawing) are unchanged by the move', rows.filter((r) => !r.sameTf));
  const meas = rows.filter((r) => r.big);
  ok(meas.length >= 8 && Math.max(...meas.map((r) => r.worst)) <= 1, `after the move the 8 handles sit on the independently computed corners and edge midpoints shifted by (37, 23) within 1 px (${meas.length} elements, worst ${Math.max(...meas.map((r) => r.worst)).toFixed(2)} px)`, meas.filter((r) => r.worst > 1));
  ok(rows.every((r) => r.dOut <= 1.01), 'the outline of every one of them moved by exactly the same 37 / 23 px (also the zero-width line, whose outline is thicker than the line)', rows.filter((r) => r.dOut > 1.01));
  ok(rows.some((r) => /rot/.test(r.name)) && rows.some((r) => /flip/.test(r.name)) && rows.some((r) => /^g0[6-9]-child|g08/.test(r.name)), 'the set includes rotated, flipped and group-nested elements (children of scaled, rotated, nested and flipped groups)');
  ok(rows.length >= 12 && (await page.evaluate(() => { const { sc } = globalThis.__s; return sc.all.every((e) => !e.world || Number.isFinite(e.world.cx)); })), 'the Scene Model still holds finite world frames for every element');
}

// ---- 5. the centre stays on the slide -----------------------------------------------------------------
{
  await draw(SEL, 0); await selectByName('front-rect');
  await page.mouse.move(350, 240); await page.mouse.down(); await page.mouse.move(2000, 1500, { steps: 6 }); await page.mouse.up();
  const o1 = (await snap()).offset;
  ok(near(o1.dx, 960 - 350, 0.1) && near(o1.dy, 540 - 240, 0.1), 'dragging far outside the slide stops with the element\'s centre on the slide edge (it can always be grabbed again)', o1);
  await page.mouse.move(2000, 1500); // pointer far outside the slide, button up
  await page.mouse.move(900, 500); await page.mouse.down(); await page.mouse.move(860, 470, { steps: 3 }); await page.mouse.move(-3000, -3000, { steps: 6 }); await page.mouse.up(); // the first part stays in the window (starts the drag and the pointer capture), the rest leaves it
  const o2 = (await snap()).offset;
  ok(near(o2.dx, -350, 0.1) && near(o2.dy, -240, 0.1), 'and the same at the top-left edge (pointer captured outside the slide, release outside the slide still ends the drag)', o2);
}

// ---- 6. cancellation and every other way a drag can end -----------------------------------------------
{
  const run = (code) => page.evaluate(code);
  await draw(SEL, 0); await selectByName('front-rect');
  const html0 = await page.evaluate(() => { globalThis.__s.sel.clear(); const h = globalThis.__s.hnd.root.outerHTML; globalThis.__s.sel.select(globalThis.__s.sc.elements.find((e) => e.name === 'front-rect').id); return h; });
  const probe = () => page.evaluate(() => { const { sel, hnd, moves } = globalThis.__s; const s = globalThis.__snap(); return { off: sel.offset(sel.id || globalThis.__s.sc.elements.find((e) => e.name === 'front-rect').id), dragging: sel.dragging, us: hnd.root.style.userSelect, moves: moves.slice(), sel: sel.id }; });
  const begin = (x = 350, y = 240, mx = 410, my = 280) => page.evaluate(({ x, y, mx, my }) => { globalThis.__s.moves.length = 0; globalThis.__ptr('pointerdown', x, y); globalThis.__ptr('pointermove', mx, my); }, { x, y, mx, my });
  const fire = (type, x, y, o) => page.evaluate(({ type, x, y, o }) => globalThis.__ptr(type, x, y, o), { type, x, y, o });

  await begin();
  let p = await probe();
  ok(p.dragging && near(p.off.dx, 60) && p.us === 'none', 'synthetic mouse drag: the element is dragged and text selection is off', p);
  await fire('pointercancel', 410, 280);
  p = await probe();
  ok(!p.dragging && p.off.dx === 0 && p.off.dy === 0 && p.us === '' && JSON.stringify(p.moves) === '["start","cancel"]', 'pointercancel puts the element back exactly where it was and reports cancel', p);
  await fire('pointerup', 410, 280);
  ok((await probe()).off.dx === 0, 'a pointerup arriving after the cancel changes nothing');

  // a kept move, then a cancelled one: back to the KEPT position, not to zero
  await begin(); await fire('pointerup', 410, 280); p = await probe();
  ok(near(p.off.dx, 60) && JSON.stringify(p.moves) === '["start","end"]', 'a release keeps the move (60 / 40)', p);
  await begin(450, 280, 500, 300); p = await probe();
  ok(near(p.off.dx, 110), 'a second drag continues from the moved position (60 + 50)', p);
  await fire('pointercancel', 500, 300); p = await probe();
  ok(near(p.off.dx, 60) && near(p.off.dy, 40), 'cancelling the second drag returns to the first move (60 / 40), not to the original place', p);
  await run(() => globalThis.__s.sel.reset());

  // Escape
  await begin(); await page.keyboard.press('Escape'); p = await probe();
  ok(!p.dragging && p.off.dx === 0 && JSON.stringify(p.moves) === '["start","cancel"]', 'Escape during the drag cancels it', p);
  await page.keyboard.press('Escape');
  ok((await probe()).off.dx === 0 && !(await probe()).dragging, 'Escape without a drag does nothing');
  // keys do not leak: the drag listeners are gone again
  const leaked = await page.evaluate(() => { let n = 0; const f = () => n++; document.addEventListener('keydown', f); document.dispatchEvent(new KeyboardEvent('keydown', { key: 'Escape', bubbles: true })); document.removeEventListener('keydown', f); return n; });
  ok(leaked === 1, 'no keyboard or selection handler is left on the document after a drag ends');

  // a second finger, a right click, another pointer
  await begin(); await fire('pointerdown', 100, 100, { isPrimary: false, pointerId: 2, pointerType: 'touch', buttons: 1 }); p = await probe();
  ok(!p.dragging && p.off.dx === 0, 'a second finger touching down during the drag cancels it (a pinch is not a move)', p);
  await fire('pointerup', 410, 280); ok((await probe()).off.dx === 0, 'and the first finger\'s release afterwards does not move or select anything');
  await begin(); await fire('pointerdown', 380, 260, { button: 2, buttons: 2 }); p = await probe();
  ok(!p.dragging && p.off.dx === 0, 'a right-button press during the drag cancels it', p);
  await begin(); await fire('pointerup', 500, 300, { pointerId: 9, isPrimary: false }); p = await probe();
  ok(p.dragging && near(p.off.dx, 60), 'a release of some other pointer does not end the drag', p);
  await fire('pointercancel', 410, 280);

  // the button was released where nobody saw it
  await begin(); await fire('pointermove', 450, 300, { buttons: 0 }); p = await probe();
  ok(!p.dragging && near(p.off.dx, 60) && JSON.stringify(p.moves) === '["start","end"]', 'a mouse move reporting no button down finishes the drag at the last position (the element never sticks to the pointer)', p);
  await run(() => globalThis.__s.sel.reset());

  // the window loses focus
  await begin(); await run(() => window.dispatchEvent(new Event('blur'))); p = await probe();
  ok(!p.dragging && p.off.dx === 0, 'the window losing focus during the drag cancels it', p);

  // selection API during a drag
  await begin(); await run(() => globalThis.__s.sel.select(globalThis.__s.sc.elements.find((e) => e.name === 'back-rect').id)); p = await probe();
  const moved = await run(() => { const { sel, sc } = globalThis.__s; return sel.offset(sc.elements.find((e) => e.name === 'front-rect').id); });
  ok(!p.dragging && moved.dx === 0, 'select(other) during a drag cancels the drag and puts the first element back', { p, moved });
  await selectByName('front-rect');
  await begin(); await run(() => globalThis.__s.sel.clear()); p = await probe();
  ok(!p.dragging && (await run(() => { const { sel, sc } = globalThis.__s; return sel.offset(sc.elements.find((e) => e.name === 'front-rect').id).dx; })) === 0 && (await run(() => !globalThis.__s.hnd.root.querySelector('.sf-sel'))), 'clear() during a drag cancels it');
  await selectByName('front-rect');
  await begin(); await run(() => globalThis.__s.sel.dispose());
  ok(await page.evaluate((h) => globalThis.__s.hnd.root.outerHTML === h, html0), 'dispose() during a drag leaves the slide byte-identical to how it was drawn');
  await fire('pointermove', 500, 300); await fire('pointerup', 500, 300);
  ok(await page.evaluate((h) => globalThis.__s.hnd.root.outerHTML === h, html0), 'and a disposed controller no longer reacts to pointer events');

  // release far from the press with no move event in between (a fast flick)
  await draw(SEL, 0); await selectByName('front-rect');
  await fire('pointerdown', 350, 240); await fire('pointerup', 450, 340); p = await probe();
  ok(!p.dragging && near(p.off.dx, 100) && near(p.off.dy, 100), 'a release far from the press without any move in between is still a drag of that distance', p);

  // touch limits: 10 px is still a tap, 12 px is a drag
  await draw(SEL, 0); await selectByName('front-rect');
  const T = { pointerType: 'touch' };
  await fire('pointerdown', 350, 240, T); await fire('pointermove', 358, 240, { ...T, buttons: 1 }); await fire('pointerup', 358, 240, T);
  ok((await probe()).off.dx === 0, 'touch: a finger that wanders 8 px (limit 10) is a tap, not a drag');
  await fire('pointerdown', 350, 240, T); await fire('pointermove', 362, 240, { ...T, buttons: 1 }); await fire('pointerup', 362, 240, T);
  ok(near((await probe()).off.dx, 12), 'touch: 12 px is a drag and the element goes the full 12 px');

  // a drag may not start while the slide is not showing (no size): no throw, no movement
  await draw(SEL, 0); await selectByName('front-rect');
  const hidden = await page.evaluate(() => { const h = document.getElementById('sf-host'); h.style.display = 'none'; try { const R = globalThis.__s.hnd.root; R.dispatchEvent(new PointerEvent('pointerdown', { bubbles: true, pointerType: 'mouse', isPrimary: true, pointerId: 1, button: 0, buttons: 1, clientX: 350, clientY: 240 })); R.dispatchEvent(new PointerEvent('pointermove', { bubbles: true, pointerType: 'mouse', isPrimary: true, pointerId: 1, buttons: 1, clientX: 450, clientY: 300 })); R.dispatchEvent(new PointerEvent('pointerup', { bubbles: true, pointerType: 'mouse', isPrimary: true, pointerId: 1, buttons: 0, clientX: 450, clientY: 300 })); return globalThis.__s.sel.offset(globalThis.__s.sel.id); } catch (e) { return 'threw ' + e; } finally { h.style.display = ''; } });
  ok(hidden && hidden.dx === 0 && hidden.dy === 0, 'a drag on a slide that is not showing (display: none) is safe and moves nothing', hidden);
}

// ---- 7. hairlines: hit-testing follows a moved line --------------------------------------------------
{
  await draw(SEL, 0); await selectByName('h-line');
  await page.mouse.move(200, 400); await page.mouse.down(); await page.mouse.move(200, 360, { steps: 5 }); await page.mouse.up();
  const pk = await page.evaluate(() => { const { sel, sc } = globalThis.__s; const n = (x, y, t) => { const id = sel.pick(x, y, t); return id ? sc.byId.get(id).name : null; }; return { off: sel.offset(sel.id), onNew: n(200, 360), twoOff: n(200, 362), farOld: n(200, 400), touchNear: n(200, 367, 'touch') }; });
  ok(near(pk.off.dy, -40) && pk.onNew === 'h-line' && pk.twoOff === 'h-line' && pk.farOld !== 'h-line' && pk.touchNear === 'h-line', 'a moved 1 pt line is hit (with the usual mouse / finger reach) at its new place and no longer at the old one', pk);
  const hs = await snap();
  ok(hs.out && near(hs.out.t + (hs.out.b - hs.out.t) / 2, 360, 1), 'its outline is centred on the moved line', hs.out);
}

// ---- 8. every fixture and real deck: a drag leaves no trace once the controller is gone ----------------
{
  const decks = [...['p31-geometry', 'p31-shapes', 'p31-text', 'p31-arabic', 'p31-media', 'p31-unsupported', 'p31-template-decor', 'p321-selection', 'p22-geometry'].map((n) => path.join(FX, n + '.pptx')), ...realFiles];
  for (const f of decks) {
    const r = await page.evaluate(async ({ b }) => {
      const { parsed, scene } = await globalThis.__load(b);
      const rn = await import('/js/render/render.js'); const sl = await import('/js/render/selection.js');
      const ser = (d) => new XMLSerializer().serializeToString(d);
      const out = { slides: 0, dragged: 0, dirty: 0, mutated: 0, worldChanged: 0, strayBefore: 0, errors: [] };
      for (let i = 0; i < parsed.slides.length; i++) {
        try {
          const sc = await scene.buildScene(parsed, i);
          const docs = [...new Set(sc.all.map((e) => e.el.ownerDocument))]; const before = docs.map(ser); const w0 = JSON.stringify(sc.all.map((e) => e.world));
          const hnd = await rn.renderSlide(sc, document.getElementById('sf-host'), { width: 960 }); const html0 = hnd.root.outerHTML;
          const sel = sl.attachSelection(hnd, sc, {});
          const R = hnd.root.getBoundingClientRect(); const ids = sc.elements.filter((e) => e.kind !== 'group' && e.kind !== 'alt' && hnd.nodes.has(e.id)).map((e) => e.id);
          let spot = null;
          for (const id of ids) { if (!sel.select(id)) continue; const r = hnd.nodes.get(id).getBoundingClientRect(); for (let a = 1; a < 10 && !spot; a++) for (let c = 1; c < 10 && !spot; c++) { const x = r.left + (r.width * a) / 10; const y = r.top + (r.height * c) / 10; if (sel.pick(x, y) === id) spot = { id, x: x - R.left, y: y - R.top }; } if (spot) break; }
          if (spot) {
            const ev = (t, x, y) => hnd.root.dispatchEvent(new PointerEvent(t, { bubbles: true, cancelable: true, pointerType: 'mouse', isPrimary: true, pointerId: 1, button: 0, buttons: t === 'pointerup' ? 0 : 1, clientX: R.left + x, clientY: R.top + y }));
            ev('pointerdown', spot.x, spot.y); ev('pointermove', spot.x + 21, spot.y + 13); ev('pointerup', spot.x + 21, spot.y + 13);
            const o = sel.offset(spot.id); if (o.dx || o.dy) out.dragged++;
          }
          sel.dispose();
          if (hnd.root.outerHTML !== html0) out.dirty++;
          docs.forEach((d, j) => { if (ser(d) !== before[j]) out.mutated++; });
          if (JSON.stringify(sc.all.map((e) => e.world)) !== w0) out.worldChanged++;
          out.slides++;
        } catch (e) { out.errors.push(i + ': ' + String(e).slice(0, 200)); }
      }
      return out;
    }, { b: b64(f) });
    const nm = path.basename(f);
    ok(r.errors.length === 0 && r.slides > 0 && r.dragged > 0, `${nm}: ${r.dragged} of ${r.slides} slide(s) took a drag without errors`, r);
    ok(r.dirty === 0 && r.mutated === 0 && r.worldChanged === 0, `${nm}: after a drag and dispose the drawn DOM is byte-identical, the XML is identical and no world frame changed`, r);
  }
}

// ---- 9. code guards ---------------------------------------------------------------------------------
{
  const code = readFileSync(path.join(ROOT, 'js/render/selection.js'), 'utf8').replace(/\/\/.*$/gm, '');
  ok(!/\.world(\.\w+)?\s*[-+*/]?=[^=]/.test(code) && !/byId\.(set|delete|clear)|scene\.(all|elements|unplaced)\.(push|splice|pop|shift)|\.rot\s*=[^=]/.test(code), 'selection.js never assigns to the Scene Model (world frames, ids, element lists)');
  ok(/setProperty\('translate'/.test(code) && !/nodes\.get\([^)]*\)\.style\.(left|top|width|height|transform)\s*=/.test(code), 'a move is only ever a `translate` on the drawn element: its left / top / size / transform are not written');
  ok(!/innerHTML|cssText|setAttribute\(\s*['"]style|localStorage|indexedDB|\bfetch\(/.test(code), 'the move code adds no innerHTML, inline style attribute, storage or network access');
}

// ======================================================================================================
section('move preview in the real UI: mouse (Rendered tab)');
{
  const ui = await ctx.newPage(); watch(ui, 'move-ui: ');
  await ui.setViewportSize({ width: 1280, height: 900 });
  await openRendered(ui, 'Move deck');
  const mv = () => ui.evaluate(() => {
    const root = document.querySelector('.rendered-host .sf-slide'); const R = root.getBoundingClientRect(); const o = root.querySelector('.sf-sel');
    const node = o && [...root.querySelectorAll('.sf-el')].find((n) => n.dataset.id === o.dataset.for);
    const rc = (n) => { const r = n.getBoundingClientRect(); return { l: r.left - R.left, t: r.top - R.top, r: r.right - R.left, b: r.bottom - R.top }; };
    const hs = [...root.querySelectorAll('.sf-sel .sf-hdl')].map((h) => { const r = h.getBoundingClientRect(); return [r.left + r.width / 2 - R.left, r.top + r.height / 2 - R.top]; });
    return { name: node?.dataset.name || null, node: node && rc(node), out: o && rc(o), hs, moved: [...root.querySelectorAll('.sf-el')].filter((n) => n.style.translate).length, status: document.querySelector('.sf-sel-status')?.textContent || '', sel: getSelection().toString(), count: root.querySelectorAll('.sf-sel').length, us: root.style.userSelect, scrollY: window.scrollY };
  });
  const drag = async (from, to, steps = 6) => { const [x0, y0] = await at(ui, ...from); const [x1, y1] = await at(ui, ...to); await ui.mouse.move(x0, y0); await ui.mouse.down(); await ui.mouse.move(x1, y1, { steps }); await ui.mouse.up(); };
  const sc = (await slideBox(ui)).s;
  await ui.mouse.click(...(await at(ui, 450, 300)));
  const a = await mv();
  ok(a.name === 'front-rect' && a.moved === 0 && /Click or tap an element|Selected: Shape/.test(a.status), 'selected front-rect, nothing moved yet', a);
  await drag([450, 300], [510, 340]);
  const b = await mv();
  ok(b.name === 'front-rect' && Math.abs(b.node.l - a.node.l - 60 * sc) <= 0.8 && Math.abs(b.node.t - a.node.t - 40 * sc) <= 0.8, `dragging the selected element with the real mouse moves it exactly (60 pt right and 40 pt down = ${(60 * sc).toFixed(1)} / ${(40 * sc).toFixed(1)} px on screen at this zoom)`, { a: a.node, b: b.node, sc });
  ok(Math.max(Math.abs(b.out.l - b.node.l), Math.abs(b.out.r - b.node.r), Math.abs(b.out.t - b.node.t), Math.abs(b.out.b - b.node.b)) <= 1 && b.hs.every((p, i) => Math.abs(p[0] - a.hs[i][0] - 60 * sc) <= 0.8 && Math.abs(p[1] - a.hs[i][1] - 40 * sc) <= 0.8), 'the outline and all handles moved with it', b);
  ok(/Selected: Shape “front-rect” \(ppt\/slides\/slide1\.xml#3\)\. Moved in the preview by 60 pt, 40 pt \(not saved\)/.test(b.status), 'the status line reports the preview move in points and says it is not saved', b.status);
  ok(b.us === '' && b.count === 1 && b.scrollY === a.scrollY, 'no text selection was left switched off, one outline, the page did not scroll');
  ok(await ui.evaluate(() => !document.querySelector('.rendered-host input, .rendered-host textarea, .rendered-host [contenteditable], .rendered-host button[type=submit]') && !document.querySelector('.sf-sel-status button, .sf-frame button')), 'still no inputs, buttons or edit / save controls on the slide or in its status line');
  // a tap on the moved element keeps it selected; the moved element is what is under the pointer now
  await ui.mouse.click(...(await at(ui, 510, 340)));
  ok((await mv()).name === 'front-rect' && (await mv()).moved === 1, 'clicking the moved element at its new place selects it (and does not undo the move)');
  await ui.mouse.click(...(await at(ui, 230, 300)));
  ok((await mv()).name !== 'front-rect', 'the old place no longer belongs to it');
  // Escape cancels the drag in the real UI too
  await ui.mouse.click(...(await at(ui, 510, 340)));
  const [ex, ey] = await at(ui, 510, 340); const [ex1, ey1] = await at(ui, 600, 380);
  const before = await mv();
  await ui.mouse.move(ex, ey); await ui.mouse.down(); await ui.mouse.move(ex1, ey1, { steps: 5 });
  const during = await mv();
  await ui.keyboard.press('Escape'); await ui.mouse.up();
  const after = await mv();
  ok(Math.abs(during.node.l - before.node.l - 90 * sc) <= 0.8 && Math.abs(after.node.l - before.node.l) <= 0.01 && Math.abs(after.node.t - before.node.t) <= 0.01, 'Escape while dragging puts the element back where it was before this drag', { before: before.node, during: during.node, after: after.node });
  await ui.mouse.move(5, 5);
  // Phase 3.2.3-B: a completed move is kept in the temporary edit-state, so it survives switching slides: coming back draws the element where it was left (selection still ends with the slide)
  const kept = await mv();
  await ui.click('.rendered .film-item >> nth=1'); await ui.waitForSelector('.rendered-host .sf-slide[data-slide="1"]');
  await ui.click('.rendered .film-item >> nth=0'); await ui.waitForSelector('.rendered-host .sf-slide[data-slide="0"]');
  const back = await mv();
  ok(back.moved === 1 && back.count === 0, 'switching away and back draws the moved element where it was left (the selection itself still ends with the slide)', back);
  await ui.mouse.click(...(await at(ui, 510, 340)));
  const backSel = await mv();
  ok(backSel.name === 'front-rect' && Math.abs(backSel.node.l - kept.node.l) <= 0.01 && Math.abs(backSel.node.t - kept.node.t) <= 0.01 && Math.abs(backSel.out.l - backSel.node.l) <= 1, 'selecting it again puts the outline on the kept position, exactly where it was before leaving the slide', { kept: kept.node, backSel: backSel.node });
  ok(/Moved in the preview by 60 pt, 40 pt \(not saved\)/.test(backSel.status), 'and the status line still reports the kept move', backSel.status);
  // the same after leaving the Rendered tab and returning, and a second move continues from the kept position
  await drag([510, 340], [560, 370]);
  const m2 = await mv();
  ok(m2.moved === 1 && Math.abs(m2.node.l - kept.node.l - 50 * sc) <= 0.8 && /by 110 pt, 70 pt/.test(m2.status), 'a second drag continues from the kept position (60 + 50 pt, 40 + 30 pt)', { kept: kept.node, m2: m2.node, status: m2.status });
  await ui.click('role=tab[name="Slides"]'); await ui.click('role=tab[name="Rendered"]'); await ui.waitForSelector('.rendered-host .sf-slide .sf-el');
  const re = await mv();
  ok(re.moved === 1 && re.name === 'front-rect' && Math.abs(re.node.l - m2.node.l) <= 0.01 && Math.abs(re.node.t - m2.node.t) <= 0.01 && Math.abs(re.out.l - re.node.l) <= 1, 'leaving the Rendered tab and coming back: the element is selected again at its kept position (outline included)', { m2: m2.node, re: re.node });
  ok(/by 110 pt, 70 pt \(not saved\)/.test(re.status), 'and the status line reports the kept move after the tab switch', re.status);
  // text: dragging a selected text box moves it and does not leave text selected
  await ui.mouse.click(...(await at(ui, 90, 320)));
  ok((await mv()).name === 'textbox', 'the text box is selected');
  await drag([90, 320], [150, 360]);
  const tx = await mv();
  ok(tx.name === 'textbox' && tx.moved === 2 && tx.sel === '', 'dragging a selected text box moves it and selects no text (front-rect\'s kept move is still there: two moved elements)', tx);
  await ui.mouse.dblclick(...(await at(ui, 150, 360)));
  ok(/Label/.test(await ui.evaluate(() => getSelection().toString())), 'a double click (a press without movement) on the selected, moved text still selects a word: only a drag moves the element');
  await ui.close();
}

// ======================================================================================================
section('move preview in the real UI: touch (phone width, real touch input)');
{
  const tctx = await browser.newContext({ viewport: { width: 1280, height: 900 }, hasTouch: true, deviceScaleFactor: 2 });
  const ui = await tctx.newPage(); watch(ui, 'move-touch: ');
  await openRendered(ui, 'Move deck');
  await ui.setViewportSize({ width: 390, height: 844 });
  await ui.waitForFunction(() => document.querySelector('.rendered-host .sf-slide').getBoundingClientRect().width < 380);
  await ui.evaluate(() => document.querySelector('.rendered-host .sf-slide').scrollIntoView({ block: 'center' }));
  const cdp = await tctx.newCDPSession(ui);
  const tap = async (x, y) => { const [px, py] = await at(ui, x, y); await ui.touchscreen.tap(px, py); };
  const info = () => ui.evaluate(() => {
    const root = document.querySelector('.rendered-host .sf-slide'); const R = root.getBoundingClientRect(); const o = root.querySelector('.sf-sel');
    const node = o && [...root.querySelectorAll('.sf-el')].find((n) => n.dataset.id === o.dataset.for); const r = node?.getBoundingClientRect(); const q = o?.getBoundingClientRect();
    return { name: node?.dataset.name || null, l: r && r.left - R.left, t: r && r.top - R.top, outDx: r && Math.max(Math.abs(q.left - r.left), Math.abs(q.top - r.top), Math.abs(q.right - r.right), Math.abs(q.bottom - r.bottom)), ta: node && getComputedStyle(node).touchAction, rootTa: getComputedStyle(root).touchAction, moved: [...root.querySelectorAll('.sf-el')].filter((n) => n.style.translate).length, scrollY: window.scrollY, count: root.querySelectorAll('.sf-sel').length, status: document.querySelector('.sf-sel-status')?.textContent || '' };
  });
  const sc = (await slideBox(ui)).s;
  const touch = async (type, x, y, extra = []) => { const [px, py] = await at(ui, x, y); await cdp.send('Input.dispatchTouchEvent', { type, touchPoints: type === 'touchEnd' ? [] : [{ x: px, y: py, id: 1 }, ...extra] }); };
  await tap(450, 300);
  const a = await info();
  ok(a.name === 'front-rect' && a.ta === 'none' && a.rootTa === 'auto', 'the selected element has touch-action: none (a finger drag on it is a move), the slide keeps auto (scroll and pinch still work)', a);
  await tap(100, 100);
  const o = await info();
  ok(o.name === 'back-rect' && o.ta === 'none', 'selecting another element moves the touch-action: none to it ...', o);
  ok(await ui.evaluate(() => [...document.querySelectorAll('.rendered-host .sf-el')].filter((n) => n.style.touchAction).length === 1), '... and leaves it on that one element only');
  await tap(450, 300);
  const f0 = await info();
  await touch('touchStart', 450, 300); await touch('touchMove', 470, 310); await touch('touchMove', 500, 330); await touch('touchMove', 510, 340); await touch('touchEnd', 510, 340);
  await ui.waitForTimeout(150);
  const f1 = await info();
  ok(f1.name === 'front-rect' && Math.abs(f1.l - f0.l - 60 * sc) <= 1 && Math.abs(f1.t - f0.t - 40 * sc) <= 1, `a finger drag on the selected element moves it 60 pt / 40 pt (${(60 * sc).toFixed(1)} / ${(40 * sc).toFixed(1)} px at this zoom)`, { f0, f1, sc });
  ok(f1.outDx <= 1 && f1.scrollY === f0.scrollY && f1.count === 1, 'the outline follows, the page did not scroll, still one selection', f1);
  ok(/Moved in the preview by 60 pt, 40 pt \(not saved\)/.test(f1.status), 'the status line reports the move', f1.status);
  // a finger drag that starts on an element that is NOT selected moves nothing
  await touch('touchStart', 100, 100); await touch('touchMove', 100, 60); await touch('touchEnd', 100, 60);
  await ui.waitForTimeout(150);
  const f2 = await info();
  ok(f2.moved === 1 && f2.name === 'front-rect', 'a swipe that starts on another element moves nothing and does not change the selection', f2);
  // a second finger during the drag cancels it
  await tap(510 + 0, 340); // the moved element
  const g0 = await info();
  const [bx, by] = await at(ui, 100, 100);
  await touch('touchStart', 510, 340); await touch('touchMove', 540, 360); await touch('touchMove', 560, 380);
  const gm = await info();
  await touch('touchMove', 560, 380, [{ x: bx, y: by, id: 2 }]);
  await touch('touchEnd', 560, 380);
  await ui.waitForTimeout(150);
  const g1 = await info();
  ok(gm.l - g0.l > 30 * sc && Math.abs(g1.l - g0.l) <= 1 && Math.abs(g1.t - g0.t) <= 1, 'a second finger touching down during the drag cancels it: the element goes back to where it was before this drag', { g0: g0.l, gm: gm.l, g1: g1.l });
  await tctx.close();
}


// ======================================================================================================
section('temporary edit-state (Phase 3.2.3-B): committed moves, group-local coordinates, redraws, cancellation, untouched XML');
const GEO = path.join(FX, 'p31-geometry.pptx');
await page.evaluate(() => {
  globalThis.__newE = async () => { const m = await import('/js/model/edit-state.js'); globalThis.__E = m.createEditState(); return true; };
  const sha = async (u8) => [...new Uint8Array(await crypto.subtle.digest('SHA-256', u8))].map((b) => b.toString(16).padStart(2, '0')).join('');
  /** One line per part of a zip: name + SHA-256 of its bytes. */
  globalThis.__zipKey = async (zip) => { const out = []; for (const n of Object.keys(zip.files).sort()) { const f = zip.files[n]; out.push(`${n}:${f.dir ? 'dir' : await sha(await f.async('uint8array'))}`); } return out.join('|'); };
  globalThis.__zipKeyOf = async (b64) => globalThis.__zipKey(await JSZip.loadAsync(Uint8Array.from(atob(b64), (c) => c.charCodeAt(0))));
  /** Name -> { rect relative to the slide, computed transforms } of every drawn slide element. */
  globalThis.__rects = () => {
    const { hnd, sc } = globalThis.__s; const R = hnd.root.getBoundingClientRect(); const out = {};
    for (const e of sc.elements) { const n = hnd.nodes.get(e.id); if (!n) continue; const r = n.getBoundingClientRect(); const geo = n.querySelector('.sf-geo'); out[e.name] = { l: r.left - R.left, t: r.top - R.top, w: r.width, h: r.height, tf: getComputedStyle(n).transform, geoTf: geo ? getComputedStyle(geo).transform : '', tr: n.style.translate || '' }; }
    return out;
  };
  /**
   * The independent check of the group-local coordinates: copy the deck IN MEMORY, write every committed element's `position` into its own <a:off> in that copy, build the Scene Model of
   * the copy with the normal (forward) code and compare every element's world corners with the corners of the untouched deck + the slide move. The original file is only read.
   */
  globalThis.__verifyLocal = async (b64, idx) => {
    const { parsed, scene, parser } = await globalThis.__load(b64);
    const sc0 = await scene.buildScene(parsed, idx);
    const recs = globalThis.__E.list().filter((r) => r.part === sc0.part);
    const zip = await JSZip.loadAsync(Uint8Array.from(atob(b64), (c) => c.charCodeAt(0)));
    const doc = new DOMParser().parseFromString(await zip.file(sc0.part).async('string'), 'application/xml');
    const byXml = new Map([...doc.getElementsByTagNameNS('*', 'cNvPr')].map((c) => [c.getAttribute('id'), c]));
    for (const r of recs) {
      const e = sc0.byId.get(r.id); const off = byXml.get(e.xmlId).parentNode.parentNode.getElementsByTagNameNS('*', 'off')[0];
      off.setAttribute('x', String(r.position.x)); off.setAttribute('y', String(r.position.y));
    }
    zip.file(sc0.part, new XMLSerializer().serializeToString(doc));
    const sc1 = await scene.buildScene(await parser.parsePptx(await zip.generateAsync({ type: 'uint8array' })), idx);
    const rows = [];
    for (const e of sc0.elements) {
      const e1 = sc1.byId.get(e.id); const r = recs.find((x) => x.id === e.id); const d = r ? r.slideDelta : { dx: 0, dy: 0 };
      let err = 0; if (e.world && e1?.world) for (let k = 0; k < 8; k++) err = Math.max(err, Math.abs(e1.world.corners[k] - (e.world.corners[k] + (k % 2 ? d.dy : d.dx) * 12700))); else err = Infinity;
      rows.push({ name: e.name, kind: e.kind, depth: e.depth, moved: !!r, err, src: e.posSource, local: r && r.localDelta, slide: r && r.slideDelta, parent: r && r.parentId });
    }
    return { rows, edited: recs.length };
  };
});

// ---- 1. every grabbable element of p31-geometry (plain, rotated, flipped, nested groups, lines, text), dragged with the real mouse ------------
const DX = 37; const DY = 23;
let geoRows = []; let geoRects0 = null; let geoSkipped = [];
{
  await page.evaluate(() => globalThis.__newE());
  await draw(GEO, 0, true); geoRects0 = await page.evaluate(() => globalThis.__rects());
  const html0 = await page.evaluate(() => globalThis.__s.hnd.root.outerHTML);
  const names = await page.evaluate(() => globalThis.__s.sc.elements.filter((e) => e.kind !== 'group' && e.kind !== 'alt').map((e) => e.name));
  for (const name of names) {
    const spot = await page.evaluate((name) => {
      const { sel, sc, hnd } = globalThis.__s; const e = sc.elements.find((x) => x.name === name); sel.select(e.id);
      const R = hnd.root.getBoundingClientRect(); const r = hnd.nodes.get(e.id).getBoundingClientRect(); let best = null;
      for (let i = 1; i < 20; i++) for (let j = 1; j < 20; j++) { const x = r.left + (r.width * i) / 20; const y = r.top + (r.height * j) / 20; if (sel.pick(x, y) === e.id) { const d = Math.hypot(x - (r.left + r.width / 2), y - (r.top + r.height / 2)); if (!best || d < best.d) best = { x: x - R.left, y: y - R.top, d }; } }
      return best;
    }, name);
    if (!spot) { geoSkipped.push(name); continue; }
    await page.mouse.move(spot.x, spot.y); await page.mouse.down(); await page.mouse.move(spot.x + DX, spot.y + DY, { steps: 5 }); await page.mouse.up();
  }
  const recs = await page.evaluate(() => globalThis.__E.list());
  ok(recs.length >= 12 && recs.length === names.length - geoSkipped.length, `each completed drag was committed: ${recs.length} elements of p31-geometry hold a move in the edit-state (one record per element)`, { recs: recs.length, skipped: geoSkipped });
  ok(recs.every((r) => Math.abs(r.slideDelta.dx - DX) < 1e-4 && Math.abs(r.slideDelta.dy - DY) < 1e-4), `every record holds the slide-space move in points: ${DX} / ${DY} (within 0.0001 pt: Chromium keeps fractional pointer coordinates as float32)`, recs.filter((r) => Math.abs(r.slideDelta.dx - DX) >= 1e-4 || Math.abs(r.slideDelta.dy - DY) >= 1e-4).map((r) => [r.id, r.slideDelta]));
  ok(recs.every((r) => JSON.stringify(Object.keys(r).sort()) === JSON.stringify(['id', 'kind', 'localDelta', 'origin', 'part', 'parentId', 'position', 'posSource', 'slideDelta', 'slideIndex'].sort())) && recs.every((r) => JSON.parse(JSON.stringify(r)) && Object.values(r).every((v) => typeof v !== 'function')), 'a record is plain data (id, part, slide index, parent id, slide delta, local delta, origin, position): no DOM or Scene Model references');
  await page.evaluate(() => globalThis.__s.sel.dispose());
  ok(await page.evaluate((h) => globalThis.__s.hnd.root.outerHTML === h, html0), 'dispose() takes every preview move off the drawn slide (byte-identical to the unmoved slide) while the edit-state keeps all of them');
  ok((await page.evaluate(() => globalThis.__E.size)) === recs.length, 'and the committed moves are still all there after the controller is gone');
}

// ---- 2. group-local coordinates: against hand-computed values and against a rebuilt Scene Model --------------------------------------------
{
  const v = await page.evaluate((b) => globalThis.__verifyLocal(b, 0), b64(GEO));
  const by = Object.fromEntries(v.rows.map((r) => [r.name, r]));
  const rad = (d) => (d * Math.PI) / 180; const EMU = 12700;
  const R = (deg, x, y) => [Math.cos(rad(deg)) * x - Math.sin(rad(deg)) * y, Math.sin(rad(deg)) * x + Math.cos(rad(deg)) * y]; // clockwise on screen (y down)
  const dxE = DX * EMU; const dyE = DY * EMU; const close = (a, b, tol = 0.5) => Math.abs(a - b) <= tol; // EMU
  const L = (n) => by[n]?.local;
  ok(v.edited >= 12 && v.rows.filter((r) => r.moved).every((r) => r.src === 'own'), `${v.edited} committed elements were written into an in-memory copy of the deck at their own <a:off> (all have their own xfrm)`, v.rows.filter((r) => r.moved && r.src !== 'own'));
  ok(v.rows.filter((r) => r.moved).every((r) => r.err < 1e-3), 'writing each record\'s group-local position into a copy of the XML and rebuilding the Scene Model puts every element\'s four corners exactly (<0.001 EMU) where the drag put them: slide corners + (37, 23) pt', v.rows.filter((r) => r.moved && r.err >= 1e-3).map((r) => [r.name, r.err]));
  ok(v.rows.filter((r) => !r.moved).every((r) => r.err < 1e-3), 'and everything that was not moved (the groups themselves included) is exactly where it was', v.rows.filter((r) => !r.moved && r.err >= 1e-3).map((r) => [r.name, r.err]));
  ok(v.rows.some((r) => r.moved && r.depth >= 2) && v.rows.some((r) => r.moved && r.depth === 1), 'the set contains top-level elements, children of a group and a child nested two groups deep', v.rows.map((r) => [r.name, r.depth]));
  ok(close(L('g01-plain').dx, dxE) && close(L('g01-plain').dy, dyE) && close(L('g02-rot30').dx, dxE) && close(L('g05-flipH-rot45').dy, dyE) && by['g01-plain'].parent === null, 'a top-level element\'s local move IS the slide move (its own rotation and flips do not change it), with no parent');
  ok(close(L('g06-child1').dx, 2 * dxE) && close(L('g06-child1').dy, 2 * dyE) && close(L('g06-child2-rot20').dx, 2 * dxE), 'child of a group drawn at scale 0.5: the local move is twice the slide move (hand-computed), also for a child that is itself rotated');
  { const [ex, ey] = R(-30, dxE, dyE); ok(close(L('g07-child').dx, ex) && close(L('g07-child').dy, ey), 'child of a group rotated 30°: the local move is the slide move turned back by 30° (hand-computed)', { got: L('g07-child'), want: [ex, ey] }); }
  ok(close(L('g09-child').dx, -dxE) && close(L('g09-child').dy, dyE), 'child of a horizontally flipped group: the horizontal part of the move is mirrored (hand-computed)', L('g09-child'));
  { // outer: scale 0.5, rot 20, flipH; inner: scale 2, rot 40 -> L = R(20) F R(40), so the local move = R(-40) F R(-20) d
    const [ax, ay] = R(-20, dxE, dyE); const [ex, ey] = R(-40, -ax, ay); ok(close(L('g08-child-rot10').dx, ex) && close(L('g08-child-rot10').dy, ey) && by['g08-child-rot10'].parent === 'ppt/slides/slide1.xml#41', 'child nested in two groups (scale, rotation, flip, scale, rotation): local move = R(-40°) F R(-20°) of the slide move (hand-computed), parent is the inner group', { got: L('g08-child-rot10'), want: [ex, ey], parent: by['g08-child-rot10'].parent }); }
  const wide = await page.evaluate(() => { const r = globalThis.__E.get('ppt/slides/slide1.xml#51'); return { o: r.origin, l: r.localDelta, p: r.position }; });
  ok(close(wide.p.x, wide.o.x + wide.l.dx, 1e-6) && close(wide.p.y, wide.o.y + wide.l.dy, 1e-6) && wide.o.x === 100000 && wide.o.y === 100000, 'a record carries the file\'s own offset (100000, 100000 EMU for g09-child) and the new position = offset + local move', wide);
}

// ---- 3. pure model: accumulation, zero, refusals, singular parents ------------------------------------------------------------------------
{
  const r = await page.evaluate(async (b) => {
    const { parsed, scene } = await globalThis.__load(b); const m = await import('/js/model/edit-state.js'); const sc = await scene.buildScene(parsed, 0);
    const id = (n) => sc.elements.find((e) => e.name === n).id; const es = m.createEditState(); const out = {};
    const a = es.commit(sc, id('g06-child1'), 10, 5); out.first = { ok: a.ok, changed: a.changed, v: es.version, local: a.edit.localDelta };
    const b2 = es.commit(sc, id('g06-child1'), 30, 15); out.total = { ok: b2.ok, slide: b2.edit.slideDelta, local: b2.edit.localDelta, size: es.size, v: es.version }; // a commit is the TOTAL offset, not a step
    const same = es.commit(sc, id('g06-child1'), 30, 15); out.same = { ok: same.ok, changed: same.changed, v: es.version };
    const z = es.commit(sc, id('g06-child1'), 0, 0); out.zero = { ok: z.ok, edit: z.edit, size: es.size, v: es.version, get: es.get(id('g06-child1')) };
    out.grp = es.commit(sc, id('g06-group-scaled'), 1, 1); out.alt = es.commit(sc, 'nope', 1, 1); out.nan = es.commit(sc, id('g01-plain'), NaN, 1); out.inf = es.commit(sc, id('g01-plain'), 1, Infinity);
    out.afterRefusals = { size: es.size, v: es.version };
    // a copy of the record cannot change the model
    es.commit(sc, id('g01-plain'), 4, 4); const c = es.get(id('g01-plain')); c.slideDelta.dx = 999; c.localDelta.dx = 999; out.isolated = es.get(id('g01-plain')).slideDelta.dx;
    // singular parent: a collapsed group (zero scale) has no inverse
    const e = sc.byId.get(id('g06-child1')); const fake = { index: 0, byId: new Map([['x', { ...e, id: 'x', world: { ...e.world, matrix: [0, 0, 0, 0, 5, 5] } }]]) };
    const sg = es.commit(fake, 'x', 3, 3); out.singular = { ok: sg.ok, reason: sg.reason, has: es.has('x') };
    const unplaced = { index: 0, byId: new Map([['u', { ...e, id: 'u', world: null }]]) }; out.unplaced = es.commit(unplaced, 'u', 3, 3);
    es.discard(id('g01-plain')); out.discard = es.size; es.commit(sc, id('g01-plain'), 1, 1); es.clear(); out.clear = es.size;
    out.deck = { same: m.editStateOf(parsed) === m.editStateOf(parsed), other: m.editStateOf(parsed) !== m.editStateOf({}) };
    return out;
  }, b64(GEO));
  ok(r.first.ok && r.first.changed && r.first.v === 1 && Math.abs(r.first.local.dx - 20 * 12700) < 1e-6, 'commit(10, 5) on a child of a 0.5-scale group: local move 20 pt x 12700 EMU, version 1', r.first);
  ok(r.total.ok && r.total.slide.dx === 30 && r.total.slide.dy === 15 && Math.abs(r.total.local.dx - 60 * 12700) < 1e-6 && r.total.size === 1 && r.total.v === 2, 'a second commit replaces the first: it is the TOTAL offset from the file (30 / 15 pt), one record per element', r.total);
  ok(r.same.ok && r.same.changed === false && r.same.v === 2, 'committing the same total again changes nothing (version unchanged)', r.same);
  ok(r.zero.ok && r.zero.edit === null && r.zero.size === 0 && r.zero.get === null, 'committing a zero offset removes the record: the element is back at its file position', r.zero);
  ok(r.grp.ok === false && r.grp.reason === 'not-movable' && r.alt.ok === false && r.alt.reason === 'not-movable' && r.nan.reason === 'invalid' && r.inf.reason === 'invalid', 'groups, unknown ids and non-finite offsets are refused', { grp: r.grp, alt: r.alt, nan: r.nan, inf: r.inf });
  ok(r.afterRefusals.size === 0 && r.afterRefusals.v === 3, 'a refusal writes nothing', r.afterRefusals);
  ok(r.isolated === 4, 'records are handed out as copies: changing one does not change the edit-state', r.isolated);
  ok(r.singular.ok === false && r.singular.reason === 'singular' && r.singular.has === false && r.unplaced.ok === false && r.unplaced.reason === 'unplaced', 'a parent that cannot be inverted (collapsed group) and an element the scene could not place are refused, nothing is stored', { s: r.singular, u: r.unplaced });
  ok(r.discard === 0 && r.clear === 0 && r.deck.same && r.deck.other, 'discard() and clear() forget moves; editStateOf(deck) is one edit-state per opened deck', r);
}

// ---- 4. redraws: the same moves, drawn again from a rebuilt scene ----------------------------------------------------------------------------
{
  const rz = () => page.evaluate(() => new Promise((r) => requestAnimationFrame(() => requestAnimationFrame(r))));
  await draw(GEO, 0, true);   // a brand-new parse, scene and drawing, the edit-state of section 1 still holds its moves
  const re = await page.evaluate(() => ({ rects: globalThis.__rects(), sel: globalThis.__s.sel.id, size: globalThis.__E.size, outlines: globalThis.__s.hnd.root.querySelectorAll('.sf-sel').length }));
  const moved = Object.keys(re.rects).filter((n) => re.rects[n].tr);
  ok(re.size >= 12 && moved.length === re.size, `a redrawn slide draws every one of the ${re.size} committed moves before anything is selected, and only those`, { moved: moved.length, size: re.size });
  ok(re.sel === null && re.outlines === 0, 'a redraw starts with no selection (the move is kept, the selection is not part of it)');
  const bad = moved.filter((n) => Math.abs(re.rects[n].l - geoRects0[n].l - DX) > 0.8 || Math.abs(re.rects[n].t - geoRects0[n].t - DY) > 0.8 || Math.abs(re.rects[n].w - geoRects0[n].w) > 0.01 || Math.abs(re.rects[n].h - geoRects0[n].h) > 0.01);
  ok(bad.length === 0, 'after the redraw each moved element sits exactly 37 px right / 23 px down of its file position, same size', bad.map((n) => [n, re.rects[n], geoRects0[n]]));
  const tf = moved.filter((n) => re.rects[n].tf !== geoRects0[n].tf || re.rects[n].geoTf !== geoRects0[n].geoTf);
  ok(tf.length === 0, 'rotation and flip transforms (frame and drawing) of every moved element are identical to the unmoved drawing', tf);
  const still = Object.keys(re.rects).filter((n) => !re.rects[n].tr && (Math.abs(re.rects[n].l - geoRects0[n].l) > 0.01 || Math.abs(re.rects[n].t - geoRects0[n].t) > 0.01));
  ok(still.length === 0, 'elements that were not moved (the ones that could not be grabbed) are where the file has them', still);
  // zoom: another width, the same slide position
  const zs = {};
  for (const w of [480, 1440, 960]) { await page.evaluate((w) => globalThis.__s.hnd.setWidth(w), w); await rz(); const r = await page.evaluate(() => globalThis.__rects()); const u = w / 960; zs[w] = moved.every((n) => Math.abs(r[n].l - geoRects0[n].l * u - DX * u) <= 0.8 && Math.abs(r[n].t - geoRects0[n].t * u - DY * u) <= 0.8); }
  ok(zs[480] && zs[1440] && zs[960], 'resizing the redrawn slide (50 %, 150 %, 100 %) keeps every moved element at the same place on the slide', zs);
  // select a moved element: its outline and handles start on the moved element
  await selectByName('g08-child-rot10');
  const sn = await snap();
  ok(sn.out && Math.max(...Object.values(shift(sn.node, sn.out)).map(Math.abs)) <= 1.5 && near(sn.offset.dx, DX, 1e-4) && near(sn.offset.dy, DY, 1e-4), 'selecting a restored element draws its outline on the moved element and reports the restored offset', { out: sn.out, node: sn.node, off: sn.offset });
  // a further move continues from the restored position, and a redraw keeps the TOTAL
  const spot = await page.evaluate(() => { const { sel, hnd } = globalThis.__s; const R = hnd.root.getBoundingClientRect(); const r = hnd.nodes.get(sel.id).getBoundingClientRect(); const ord = [10, 9, 11, 8, 12, 7, 13, 6, 14, 5, 15, 4, 16, 3, 17, 2, 18, 1, 19]; for (const i of ord) for (const j of ord) { const x = r.left + (r.width * i) / 20; const y = r.top + (r.height * j) / 20; if (sel.pick(x, y) === sel.id) return { x: x - R.left, y: y - R.top }; } return null; }); // from the centre outwards: a press near a corner or an edge centre is on a resize handle since Phase 3.2.4-A
  await page.mouse.move(spot.x, spot.y); await page.mouse.down(); await page.mouse.move(spot.x + 10, spot.y - 6, { steps: 4 }); await page.mouse.up();
  const e2 = await page.evaluate(() => { const { sel } = globalThis.__s; return { off: sel.offset(sel.id), rec: globalThis.__E.get(sel.id) }; });
  ok(near(e2.off.dx, DX + 10, 1e-4) && near(e2.off.dy, DY - 6, 1e-4) && near(e2.rec.slideDelta.dx, DX + 10, 1e-4) && near(e2.rec.slideDelta.dy, DY - 6, 1e-4), 'dragging a restored element goes on from where it was left (37 + 10, 23 - 6 pt) and the record holds the total', e2);
  await draw(GEO, 0, true);
  const again = await page.evaluate(() => globalThis.__s.sc.elements.map((e) => [e.name, globalThis.__E.get(e.id)?.slideDelta]).filter((x) => x[1]).find((x) => x[0] === 'g08-child-rot10'));
  ok(again && near(again[1].dx, DX + 10, 1e-4) && near(again[1].dy, DY - 6, 1e-4), 'and after another redraw the total is what is held and drawn', again);
  const v2 = await page.evaluate((b) => globalThis.__verifyLocal(b, 0), b64(GEO));
  ok(v2.rows.every((r) => r.err < 1e-3), 'the group-local coordinates still reproduce every dragged position exactly after the second move (deck copy rebuilt, all corners within 0.001 EMU)', v2.rows.filter((r) => r.err >= 1e-3).map((r) => [r.name, r.err]));
}

// ---- 5. cancellation never reaches the edit-state --------------------------------------------------------------------------------------------
{
  await page.evaluate(() => globalThis.__newE());
  await draw(SEL, 0, true); await selectByName('front-rect');
  const FR = 'ppt/slides/slide1.xml#3';
  const run = (code) => page.evaluate(code);
  const begin = (x = 350, y = 240, mx = 410, my = 280) => page.evaluate(({ x, y, mx, my }) => { globalThis.__s.moves.length = 0; globalThis.__ptr('pointerdown', x, y); globalThis.__ptr('pointermove', mx, my); }, { x, y, mx, my });
  const fire = (type, x, y, o) => page.evaluate(({ type, x, y, o }) => globalThis.__ptr(type, x, y, o), { type, x, y, o });
  const es = () => page.evaluate((id) => ({ size: globalThis.__E.size, v: globalThis.__E.version, rec: globalThis.__E.get(id), off: globalThis.__s.sel.offset(id), dragging: globalThis.__s.sel.dragging, moves: globalThis.__s.moves.slice() }), FR);
  const base = await es();
  ok(base.size === 0 && base.v === 0, 'a fresh edit-state is empty (version 0)', base);
  await begin(); let s = await es();
  ok(s.dragging === FR && Math.abs(s.off.dx - 60) < 1e-9 && s.size === 0 && s.v === 0 && s.rec === null, 'while a drag is open the edit-state is untouched (the move is only a preview until it completes)', s);
  const cancels = [
    ['pointercancel', () => fire('pointercancel', 410, 280)],
    ['Escape', () => page.keyboard.press('Escape')],
    ['the window losing focus', () => run(() => window.dispatchEvent(new Event('blur')))],
    ['a second finger', () => fire('pointerdown', 100, 100, { isPrimary: false, pointerId: 2, pointerType: 'touch', buttons: 1 })],
    ['a right-button press', () => fire('pointerdown', 380, 260, { button: 2, buttons: 2 })],
    ['select(other element)', () => run(() => globalThis.__s.sel.select(globalThis.__s.sc.elements.find((e) => e.name === 'back-rect').id))],
    ['clear()', () => run(() => globalThis.__s.sel.clear())],
    ['dispose()', () => run(() => globalThis.__s.sel.dispose())],
  ];
  for (const [label, act] of cancels) {
    if (label !== 'pointercancel') { await draw(SEL, 0, true); await selectByName('front-rect'); }
    await begin(); await act(); s = await es();
    ok(!s.dragging && s.size === 0 && s.v === 0 && s.rec === null && s.off.dx === 0, `${label} during a drag: the element is back, nothing was committed, the edit-state was not written to at all`, s);
  }
  // a kept move, then cancelled moves: the kept record is exactly what it was
  await draw(SEL, 0, true); await selectByName('front-rect');
  await begin(); await fire('pointerup', 410, 280); const kept = await es();
  ok(kept.size === 1 && kept.v === 1 && kept.rec && Math.abs(kept.rec.slideDelta.dx - 60) < 1e-9 && Math.abs(kept.rec.slideDelta.dy - 40) < 1e-9 && JSON.stringify(kept.moves) === '["start","end"]', 'a release commits the move: one record, 60 / 40 pt, version 1', kept);
  for (const [label, act] of [['pointercancel', () => fire('pointercancel', 500, 300)], ['Escape', () => page.keyboard.press('Escape')], ['a second finger', () => fire('pointerdown', 100, 100, { isPrimary: false, pointerId: 2, pointerType: 'touch', buttons: 1 })], ['select(other element)', () => run(() => globalThis.__s.sel.select(globalThis.__s.sc.elements.find((e) => e.name === 'back-rect').id))]]) {
    if (label === 'select(other element)') await selectByName('front-rect');
    await begin(450, 280, 500, 300); const mid = await es();
    await act(); s = await es();
    ok(mid.dragging === FR && Math.abs(mid.off.dx - 110) < 1e-9 && JSON.stringify(mid.rec) === JSON.stringify(kept.rec), `second drag open (110 pt): the committed record still holds the first move`, mid);
    ok(!s.dragging && JSON.stringify(s.rec) === JSON.stringify(kept.rec) && s.v === 1 && Math.abs(s.off.dx - 60) < 1e-9, `${label} cancels the second drag: the committed record is byte-identical to before, no write (version 1), the element is back at 60 pt`, s);
  }
  // a drag that returns to where it started commits nothing new (the loop above left another element selected: select the dragged one again)
  await selectByName('front-rect');
  await begin(450, 280, 500, 300); await fire('pointermove', 450, 280); await fire('pointerup', 450, 280); s = await es();
  ok(!s.dragging && JSON.stringify(s.rec) === JSON.stringify(kept.rec) && s.v === 1 && JSON.stringify(s.moves) === '["start","end"]', 'a drag that ends where it started leaves the record and the version as they were', s);
  // another pointer's events never commit: its release does not end the drag, and cancelling afterwards leaves the record alone
  await begin(450, 280, 500, 300); await fire('pointerup', 500, 300, { pointerId: 9, isPrimary: false }); s = await es();
  ok(s.dragging === FR && s.v === 1 && JSON.stringify(s.rec) === JSON.stringify(kept.rec), 'the release of some other pointer neither ends the drag nor commits', s);
  await fire('pointercancel', 500, 300); s = await es();
  ok(!s.dragging && s.v === 1 && JSON.stringify(s.rec) === JSON.stringify(kept.rec) && Math.abs(s.off.dx - 60) < 1e-9, 'and the cancel that follows puts the element back at the kept move without writing', s);
  // a flick (a release far from the press, no move event in between) commits the total
  await fire('pointerdown', 450, 280); await fire('pointerup', 500, 300); s = await es();
  ok(!s.dragging && s.v === 2 && Math.abs(s.rec.slideDelta.dx - 110) < 1e-9 && Math.abs(s.rec.slideDelta.dy - 60) < 1e-9, 'the release of the dragging pointer commits the total (110 / 60 pt), version 2', s);
  // reset() is the explicit undo: element back, record gone
  await run(() => globalThis.__s.sel.reset()); s = await es();
  ok(s.size === 0 && s.rec === null && s.off.dx === 0 && s.v === 3, 'reset() puts the element back AND forgets its committed move', s);
  // an edit-state that refuses a move: the drag counts as cancelled and nothing is left drawn
  const ref = await page.evaluate(async (b) => {
    const { parsed, scene } = await globalThis.__load(b); const rn = await import('/js/render/render.js'); const sl = await import('/js/render/selection.js');
    const sc = await scene.buildScene(parsed, 0); globalThis.__s?.sel.dispose();
    const hnd = await rn.renderSlide(sc, document.getElementById('sf-host'), { width: 960 }); const phases = []; let asked = 0;
    const sel = sl.attachSelection(hnd, sc, { onMove: (m) => phases.push(m.phase), edits: { forSlide: () => [], commit: () => { asked++; return { ok: false, reason: 'singular' }; }, discard() {} } });
    globalThis.__s = { hnd, sc, sel, parsed, changes: [], moves: [] };
    const id = sc.elements.find((e) => e.name === 'front-rect').id; sel.select(id);
    globalThis.__ptr('pointerdown', 350, 240); globalThis.__ptr('pointermove', 410, 280); globalThis.__ptr('pointerup', 410, 280);
    return { phases, asked, off: sel.offset(id), moved: [...hnd.root.querySelectorAll('.sf-el')].filter((n) => n.style.translate).length, dragging: sel.dragging };
  }, b64(SEL));
  ok(ref.asked === 1 && JSON.stringify(ref.phases) === '["start","cancel"]' && ref.off.dx === 0 && ref.moved === 0 && ref.dragging === null, 'if the edit-state refuses the move, the release counts as a cancel: the element goes back and the slide never shows a move the model does not hold', ref);
}

// ---- 6. the original XML, the Scene Model and the file are untouched --------------------------------------------------------------------------
{
  const decks = [GEO, SEL, path.join(FX, 'p31-shapes.pptx'), path.join(FX, 'p31-text.pptx'), path.join(FX, 'p31-arabic.pptx'), path.join(FX, 'p31-media.pptx'), path.join(FX, 'p31-unsupported.pptx'), path.join(FX, 'p31-template-decor.pptx'), path.join(FX, 'p22-geometry.pptx'), ...realFiles];
  for (const f of decks) {
    const r = await page.evaluate(async ({ b }) => {
      const { parsed, scene } = await globalThis.__load(b); const rn = await import('/js/render/render.js'); const sl = await import('/js/render/selection.js'); const em = await import('/js/model/edit-state.js');
      const ser = (d) => new XMLSerializer().serializeToString(d); const edits = em.createEditState();
      const zip0 = await globalThis.__zipKey(parsed.zip); const orig = await globalThis.__zipKeyOf(b);
      const out = { slides: 0, committed: 0, restored: 0, mutated: 0, worldChanged: 0, xfrmChanged: 0, dirty: 0, errors: [], zipSame: zip0 === orig, zipAfter: true };
      for (let i = 0; i < parsed.slides.length; i++) {
        try {
          const sc = await scene.buildScene(parsed, i); const docs = [...new Set(sc.all.map((e) => e.el.ownerDocument))]; const before = docs.map(ser);
          const w0 = JSON.stringify(sc.all.map((e) => [e.world, e.xfrm]));
          const hnd = await rn.renderSlide(sc, document.getElementById('sf-host'), { width: 960 }); const html0 = hnd.root.outerHTML;
          let sel = sl.attachSelection(hnd, sc, { edits }); const R = hnd.root.getBoundingClientRect();
          let spot = null;
          for (const e of sc.elements) { if (e.kind === 'group' || e.kind === 'alt' || !hnd.nodes.has(e.id) || !sel.select(e.id)) continue; const r = hnd.nodes.get(e.id).getBoundingClientRect(); for (let a = 1; a < 10 && !spot; a++) for (let c = 1; c < 10 && !spot; c++) { const x = r.left + (r.width * a) / 10; const y = r.top + (r.height * c) / 10; if (sel.pick(x, y) === e.id) spot = { id: e.id, x: x - R.left, y: y - R.top }; } if (spot) break; }
          if (spot) {
            const ev = (t, x, y) => hnd.root.dispatchEvent(new PointerEvent(t, { bubbles: true, cancelable: true, pointerType: 'mouse', isPrimary: true, pointerId: 1, button: 0, buttons: t === 'pointerup' ? 0 : 1, clientX: R.left + x, clientY: R.top + y }));
            ev('pointerdown', spot.x, spot.y); ev('pointermove', spot.x + 21, spot.y + 13); ev('pointerup', spot.x + 21, spot.y + 13);
            if (edits.has(spot.id)) out.committed++;
            sel.dispose(); if (hnd.root.outerHTML !== html0) out.dirty++;
            sel = sl.attachSelection(hnd, sc, { edits }); // redraw of the same slide: the move is drawn again
            const o = sel.offset(spot.id); const rec = edits.get(spot.id);
            if (rec && Math.abs(o.dx - rec.slideDelta.dx) < 1e-9 && Math.abs(o.dy - rec.slideDelta.dy) < 1e-9 && hnd.nodes.get(spot.id).style.translate) out.restored++;
          }
          sel.dispose(); if (hnd.root.outerHTML !== html0) out.dirty++;
          docs.forEach((d, j) => { if (ser(d) !== before[j]) out.mutated++; });
          if (JSON.stringify(sc.all.map((e) => [e.world, e.xfrm])) !== w0) out.worldChanged++;
          out.slides++;
        } catch (e) { out.errors.push(i + ': ' + String(e).slice(0, 200)); }
      }
      out.zipAfter = (await globalThis.__zipKey(parsed.zip)) === orig;
      return out;
    }, { b: b64(f) });
    const nm = path.basename(f);
    ok(r.errors.length === 0 && r.slides > 0 && r.committed > 0 && r.restored === r.committed, `${nm}: ${r.committed} of ${r.slides} slide(s) committed a drag and drew it again from the edit-state without errors`, r);
    ok(r.dirty === 0 && r.mutated === 0 && r.worldChanged === 0 && r.zipSame && r.zipAfter, `${nm}: slide / layout / master XML, every part of the file (SHA-256 per part, before and after), every world frame and xfrm are identical, and the drawn DOM is clean after dispose`, r);
  }
  // the edit-state module is data only
  const code = readFileSync(path.join(ROOT, 'js/model/edit-state.js'), 'utf8').replace(/\/\/.*$/gm, '');
  ok(!/\b(document|window|localStorage|sessionStorage|indexedDB|fetch|XMLSerializer|DOMParser|setAttribute|innerHTML|JSZip)\b/.test(code), 'edit-state.js touches no DOM, storage, network, XML serializer or zip');
  ok(!/\.(world|xfrm|el)\s*[-+*/]?=[^=]|\bdoc\b|\.ownerDocument|byId\.(set|delete|clear)/.test(code) && [...code.matchAll(/^import .* from '(.*)';$/gm)].map((m) => m[1]).join() === './scene.js', 'edit-state.js never assigns to the Scene Model or an XML node, and imports nothing but the matrix helpers of scene.js');
  const sel = readFileSync(path.join(ROOT, 'js/render/selection.js'), 'utf8').replace(/\/\/.*$/gm, '');
  ok(!/\.xfrm\b|\.localDelta\b|\.position\.[xy]\b|\.origin\.[xy]\b/.test(sel), 'selection.js hands the offset to the edit-state and never computes or writes a local coordinate itself');
  const rd = readFileSync(path.join(ROOT, 'js/ui/rendered.js'), 'utf8').replace(/\/\/.*$/gm, '');
  ok(!/localStorage|indexedDB|store\.templates|\.update\(/.test(rd), 'the Rendered tab keeps the edit-state in memory only: no storage call was added');
}

// ---- 7. in the real UI: slides, tabs, accumulation, and nothing outlives the opened deck -------------------------------------------------------
{
  const ui = await ctx.newPage(); watch(ui, 'commit-ui: ');
  await ui.setViewportSize({ width: 1280, height: 900 });
  await openRendered(ui, 'Commit deck');
  const store0 = await ui.evaluate(async () => ({ ls: JSON.stringify({ ...localStorage }), dbs: ((await indexedDB.databases?.()) || []).map((d) => d.name).sort().join() }));
  const info = () => ui.evaluate(() => { const root = document.querySelector('.rendered-host .sf-slide'); const R = root.getBoundingClientRect(); const out = {}; for (const n of root.querySelectorAll('.sf-el')) { const r = n.getBoundingClientRect(); out[n.dataset.name] = { l: r.left - R.left, t: r.top - R.top, tr: n.style.translate }; } return { els: out, status: document.querySelector('.sf-sel-status')?.textContent || '', slide: root.dataset.slide }; });
  const sc = (await slideBox(ui)).s;
  const drag = async (from, to) => { const [x0, y0] = await at(ui, ...from); const [x1, y1] = await at(ui, ...to); await ui.mouse.move(x0, y0); await ui.mouse.down(); await ui.mouse.move(x1, y1, { steps: 6 }); await ui.mouse.up(); };
  const a = await info();
  await ui.mouse.click(...(await at(ui, 450, 300))); await drag([450, 300], [510, 340]);
  const b = await info();
  const shifted = (x, y) => Math.abs(b.els['front-rect'].l - a.els['front-rect'].l - x * sc) <= 0.8 && Math.abs(b.els['front-rect'].t - a.els['front-rect'].t - y * sc) <= 0.8;
  ok(shifted(60, 40), 'real mouse: the element moved 60 / 40 pt');
  // slide 2 and back, three times, with a move on slide 2 as well
  await ui.click('.rendered .film-item >> nth=1'); await ui.waitForSelector('.rendered-host .sf-slide[data-slide="1"]');
  const s2a = await info();
  await ui.mouse.click(...(await at(ui, 150, 150))); await drag([150, 150], [200, 190]);
  const s2b = await info();
  ok(s2b.els['s2-rect'] && Math.abs(s2b.els['s2-rect'].l - s2a.els['s2-rect'].l - 50 * sc) <= 0.8, 'a move on slide 2 works the same', { s2a: s2a.els['s2-rect'], s2b: s2b.els['s2-rect'] });
  ok(Object.keys(s2b.els).filter((n) => s2b.els[n].tr).length === 1 && !s2b.els['front-rect'], 'slide 2 shows only its own move (slide 1\'s move does not leak onto it)', Object.keys(s2b.els));
  for (let k = 0; k < 3; k++) {
    await ui.click('.rendered .film-item >> nth=0'); await ui.waitForSelector('.rendered-host .sf-slide[data-slide="0"] .sf-el');
    const c1 = await info();
    ok(Math.abs(c1.els['front-rect'].l - b.els['front-rect'].l) <= 0.01 && Math.abs(c1.els['front-rect'].t - b.els['front-rect'].t) <= 0.01 && Object.keys(c1.els).filter((n) => c1.els[n].tr).join() === 'front-rect', `slide 1 again (round ${k + 1}): front-rect is exactly where it was left and it is the only moved element`, { b: b.els['front-rect'], c1: c1.els['front-rect'] });
    await ui.click('.rendered .film-item >> nth=1'); await ui.waitForSelector('.rendered-host .sf-slide[data-slide="1"] .sf-el');
    const c2 = await info();
    ok(Math.abs(c2.els['s2-rect'].l - s2b.els['s2-rect'].l) <= 0.01 && Math.abs(c2.els['s2-rect'].t - s2b.els['s2-rect'].t) <= 0.01, `slide 2 again (round ${k + 1}): s2-rect is exactly where it was left`);
  }
  // the other tabs and back, and the filmstrip of the Slides tab
  await ui.click('role=tab[name="Design"]'); await ui.click('role=tab[name="Slides"]'); await ui.click('.preview:not(.rendered) .film-item >> nth=0'); await ui.click('role=tab[name="Design"]'); await ui.click('role=tab[name="Rendered"]'); await ui.waitForSelector('.rendered-host .sf-slide[data-slide="0"] .sf-el');
  const t1 = await info();
  ok(Math.abs(t1.els['front-rect'].l - b.els['front-rect'].l) <= 0.01 && Math.abs(t1.els['front-rect'].t - b.els['front-rect'].t) <= 0.01 && t1.slide === '0', 'Slides, Design and Rendered tab round trips (and choosing the slide in the Slides tab) keep the move', { b: b.els['front-rect'], t1: t1.els['front-rect'] });
  // resize of the window after the redraw
  await ui.setViewportSize({ width: 820, height: 900 });
  await ui.waitForFunction(() => { const r = document.querySelector('.rendered-host .sf-slide'); const f = document.querySelector('.sf-frame'); return r && Math.abs(r.getBoundingClientRect().width - f.clientWidth) <= 1 && f.clientWidth < 800; });
  const t2 = await info(); const sc2 = (await slideBox(ui)).s;
  ok(Math.abs(t2.els['front-rect'].l - (a.els['front-rect'].l / sc) * sc2 - 60 * sc2) <= 1 && Math.abs(t2.els['front-rect'].t - (a.els['front-rect'].t / sc) * sc2 - 40 * sc2) <= 1, 'after the window is resized the kept move is still 60 / 40 pt from the file position (the offset is in slide points)', { t2: t2.els['front-rect'], sc2 });
  await ui.setViewportSize({ width: 1280, height: 900 });
  // nothing was stored
  const store1 = await ui.evaluate(async () => ({ ls: JSON.stringify({ ...localStorage }), dbs: ((await indexedDB.databases?.()) || []).map((d) => d.name).sort().join() }));
  ok(store1.ls === store0.ls && store1.dbs === store0.dbs, 'moving elements wrote nothing to localStorage and created no IndexedDB database', { store0, store1 });
  // closing the deck ends the edit-state: opening it again shows the file
  await ui.click('.modal [aria-label="Close"]'); await ui.waitForSelector('.modal', { state: 'detached' });
  await ui.click('.tpl:has-text("Commit deck") >> text=Preview'); await ui.waitForSelector('.film-item');
  await ui.click('role=tab[name="Rendered"]'); await ui.waitForSelector('.rendered-host .sf-slide .sf-el');
  const r1 = await info();
  ok(Object.values(r1.els).every((e) => !e.tr) && Math.abs(r1.els['front-rect'].l - a.els['front-rect'].l) <= 0.01 && Math.abs(r1.els['front-rect'].t - a.els['front-rect'].t) <= 0.01, 'closing the preview and opening the template again shows every element at its file position: the edit-state is temporary and ends with the opened deck', r1);
  await ui.close();
}

// ======================================================================================================
section('movement hardening (Phase 3.2.3-C): a stray pointerup neither drops a press nor orphans a drag');
{
  const FR = 'ppt/slides/slide1.xml#3';
  const eq = (a, b) => Math.abs(a - b) < 1e-6;
  const run = (code) => page.evaluate(code);
  const begin = (x = 350, y = 240, mx = 410, my = 280) => page.evaluate(({ x, y, mx, my }) => { globalThis.__s.moves.length = 0; globalThis.__ptr('pointerdown', x, y); globalThis.__ptr('pointermove', mx, my); }, { x, y, mx, my });
  const fire = (type, x, y, o) => page.evaluate(({ type, x, y, o }) => globalThis.__ptr(type, x, y, o), { type, x, y, o });
  const es = () => page.evaluate((id) => ({ size: globalThis.__E.size, v: globalThis.__E.version, rec: globalThis.__E.get(id), off: globalThis.__s.sel.offset(id), dragging: globalThis.__s.sel.dragging, selId: globalThis.__s.sel.id, us: globalThis.__s.hnd.root.style.userSelect, moves: globalThis.__s.moves.slice() }), FR);
  const fresh = async () => { await page.evaluate(() => globalThis.__newE()); await draw(SEL, 0, true); await selectByName('front-rect'); };
  const STRAY = [['a mouse-like pointer with another id', { pointerId: 9, isPrimary: false }], ['a second finger', { pointerId: 5, isPrimary: false, pointerType: 'touch' }], ['a pen', { pointerId: 6, isPrimary: false, pointerType: 'pen' }], ['a primary pointer with another id', { pointerId: 7 }]];

  // 1. the release of another pointer in the middle of a drag: the drag stays open, keeps following the real pointer, and the real release keeps and commits it
  for (const [label, o] of STRAY) {
    await fresh();
    await begin(); let s = await es();
    ok(s.dragging === FR && eq(s.off.dx, 60) && eq(s.off.dy, 40), `${label}: the drag is open at 60 / 40 pt`, s);
    await fire('pointerup', 500, 300, o); s = await es();
    ok(s.dragging === FR && eq(s.off.dx, 60) && s.v === 0 && s.size === 0, `${label}: its release in the middle of the drag leaves the drag open and commits nothing`, s);
    await fire('pointermove', 450, 300); s = await es();
    ok(s.dragging === FR && eq(s.off.dx, 100) && eq(s.off.dy, 60), `${label}: the element still follows the real pointer afterwards (100 / 60 pt)`, s);
    await fire('pointerup', 450, 300); s = await es();
    ok(!s.dragging && eq(s.off.dx, 100) && eq(s.off.dy, 60) && s.v === 1 && s.size === 1 && s.rec && eq(s.rec.slideDelta.dx, 100) && eq(s.rec.slideDelta.dy, 60) && JSON.stringify(s.moves) === '["start","end"]' && s.us === '', `${label}: the real release keeps the move and commits it once (100 / 60 pt, version 1); text selection is restored`, s);
  }
  // after a finished drag no handler is left on the document, and a duplicate release changes nothing
  await fire('pointerup', 450, 300); let s = await es();
  ok(!s.dragging && s.v === 1 && eq(s.off.dx, 100) && s.selId === FR, 'a duplicate release after the drag ended changes nothing (no second commit, still selected)', s);
  const leaked = await page.evaluate(() => { let n = 0; const f = () => n++; document.addEventListener('keydown', f); document.dispatchEvent(new KeyboardEvent('keydown', { key: 'Escape', bubbles: true })); document.removeEventListener('keydown', f); return n; });
  ok(leaked === 1, 'no keyboard handler is left on the document after a drag that saw a stray release');

  // 2. the release of another pointer while the press is still pending (before the travel limit): the drag can still start
  for (const [label, o] of STRAY.slice(0, 2)) {
    await fresh();
    await page.evaluate(() => { globalThis.__s.moves.length = 0; });
    await fire('pointerdown', 350, 240); await fire('pointerup', 600, 100, o); await fire('pointermove', 410, 280); s = await es();
    ok(s.dragging === FR && eq(s.off.dx, 60) && eq(s.off.dy, 40), `${label}: a stray release before the drag started does not drop the press (the drag starts, 60 / 40 pt)`, s);
    await fire('pointerup', 410, 280); s = await es();
    ok(!s.dragging && s.v === 1 && eq(s.rec.slideDelta.dx, 60) && eq(s.rec.slideDelta.dy, 40) && JSON.stringify(s.moves) === '["start","end"]', `${label}: and the real release commits it`, s);
  }

  // 3. a press that is still in flight stays a click when another pointer's release comes first
  await fresh(); await run(() => globalThis.__s.sel.clear());
  await fire('pointerdown', 350, 240); await fire('pointerup', 350, 240, { pointerId: 9, isPrimary: false }); s = await es();
  ok(s.selId === null, 'a stray release alone does not select anything');
  await fire('pointerup', 350, 240); s = await es();
  ok(s.selId === FR && !s.dragging && s.v === 0, 'the click still selects when its own release arrives after a stray one', s);
  await fire('pointerdown', 30, 30); await fire('pointerup', 30, 30, { pointerId: 9, isPrimary: false }); await fire('pointerup', 30, 30); s = await es();
  ok(s.selId === null, 'and a click on empty space still clears after a stray release', s);

  // 4. a release without any press does nothing, whoever it comes from
  await fresh();
  await fire('pointerup', 30, 30); await fire('pointerup', 30, 30, { pointerId: 9, isPrimary: false }); await fire('pointerup', 350, 240); s = await es();
  ok(s.selId === FR && !s.dragging && s.v === 0 && s.size === 0 && eq(s.off.dx, 0), 'releases without a press (own or foreign) change nothing: still selected, nothing moved or committed', s);

  // 5. a stray release does not change how a drag is cancelled
  for (const [label, act] of [['pointercancel', () => fire('pointercancel', 410, 280)], ['Escape', () => page.keyboard.press('Escape')]]) {
    await fresh(); await begin(); await fire('pointerup', 500, 300, { pointerId: 9, isPrimary: false }); await act(); s = await es();
    ok(!s.dragging && eq(s.off.dx, 0) && s.v === 0 && s.size === 0 && s.us === '' && JSON.stringify(s.moves) === '["start","cancel"]', `${label} after a stray release still puts the element back and writes nothing`, s);
  }

  // 6. a second drag after a committed move, with a stray release in the middle: 60 + 70 and 40 + 50
  await fresh(); await begin(); await fire('pointerup', 410, 280); s = await es();
  ok(s.v === 1 && eq(s.rec.slideDelta.dx, 60), 'first drag committed (60 / 40 pt, version 1)', s);
  await begin(450, 280, 500, 300); await fire('pointerup', 500, 300, { pointerId: 9, isPrimary: false }); await fire('pointermove', 520, 330); s = await es();
  ok(s.dragging === FR && eq(s.off.dx, 130) && eq(s.off.dy, 90) && s.v === 1, 'second drag after a stray release follows the pointer from the kept move (130 / 90 pt) and has not committed yet', s);
  await fire('pointerup', 520, 330); s = await es();
  ok(!s.dragging && s.v === 2 && eq(s.rec.slideDelta.dx, 130) && eq(s.rec.slideDelta.dy, 90) && s.size === 1, 'the real release commits the total (130 / 90 pt, version 2, still one record)', s);

  // 7. the real Rendered tab (events dispatched on the real slide root: a real mouse cannot produce a release of another pointer)
  {
    const ui = await ctx.newPage(); watch(ui, 'stray-ui: ');
    await ui.setViewportSize({ width: 1280, height: 900 });
    await openRendered(ui, 'Stray pointerup deck');
    const sc = (await slideBox(ui)).s;
    const fr = () => ui.evaluate(() => { const root = document.querySelector('.rendered-host .sf-slide'); const R = root.getBoundingClientRect(); const n = [...root.querySelectorAll('.sf-el')].find((x) => x.dataset.name === 'front-rect'); const r = n.getBoundingClientRect(); const o = root.querySelector('.sf-sel'); const q = o && o.getBoundingClientRect(); return { l: r.left - R.left, t: r.top - R.top, ol: q ? q.left - R.left : null, ot: q ? q.top - R.top : null, moved: [...root.querySelectorAll('.sf-el')].filter((x) => x.style.translate).length, status: document.querySelector('.sf-sel-status')?.textContent || '', us: root.style.userSelect }; });
    const send = (list) => ui.evaluate((list) => { const root = document.querySelector('.rendered-host .sf-slide'); for (const [type, x, y, o] of list) root.dispatchEvent(new PointerEvent(type, { bubbles: true, cancelable: true, composed: true, pointerType: 'mouse', isPrimary: true, pointerId: 1, button: 0, buttons: type === 'pointerup' ? 0 : 1, clientX: x, clientY: y, ...o })); }, list);
    await ui.mouse.click(...(await at(ui, 450, 300)));
    const a = await fr();
    const [p0, p1, p2] = [await at(ui, 450, 300), await at(ui, 510, 340), await at(ui, 530, 350)];
    await send([['pointerdown', ...p0], ['pointermove', ...p1], ['pointerup', ...p1, { pointerId: 9, isPrimary: false, pointerType: 'touch' }], ['pointermove', ...p2], ['pointerup', ...p2]]);
    const b = await fr();
    ok(Math.abs(b.l - a.l - 80 * sc) <= 0.8 && Math.abs(b.t - a.t - 50 * sc) <= 0.8 && Math.abs(b.ol - b.l) <= 1 && Math.abs(b.ot - b.t) <= 1, `real tab: a stray release in the middle of a drag does not stop it: the element and its outline end 80 / 50 pt from where they started (${(80 * sc).toFixed(1)} / ${(50 * sc).toFixed(1)} px)`, { a, b, sc });
    ok(/Moved in the preview by 80 pt, 50 pt \(not saved\)/.test(b.status) && b.us === '', 'real tab: the status line reports the finished move and text selection is back on', b);
    await ui.click('.rendered .film-item >> nth=1'); await ui.waitForSelector('.rendered-host .sf-slide[data-slide="1"] .sf-el');
    await ui.click('.rendered .film-item >> nth=0'); await ui.waitForSelector('.rendered-host .sf-slide[data-slide="0"] .sf-el');
    const c = await fr();
    ok(c.moved === 1 && Math.abs(c.l - b.l) <= 0.01 && Math.abs(c.t - b.t) <= 0.01, 'real tab: that move was committed (not left half-open): it is drawn again where it was left after another slide and back', { b, c });
    await ui.close();
  }

  // 8. code guard: onUp ignores a pointerup that is not the pressing pointer's before it touches the press or the drag
  const code = readFileSync(path.join(ROOT, 'js/render/selection.js'), 'utf8');
  ok(/const onUp = \(e\) => \{\s*const d = down; if \(!d \|\| e\.pointerId !== d\.id\) return;[^\n]*\n\s*down = null;/.test(code), 'selection.js: onUp returns for a foreign or unpaired pointerup before it clears the press');
}

// ======================================================================================================
section('resize preview (Phase 3.2.4-A): drag the 8 handles, temporary, preview only');
{
  await page.evaluate(() => {
    const g = globalThis;
    /** Corners of a node as drawn (slide px), from its resolved layout box and its transforms: independent of getBoundingClientRect. */
    g.__corners = (n) => {
      const cs = getComputedStyle(n);
      const W = parseFloat(cs.width); const H = parseFloat(cs.height);
      const m = new DOMMatrix(cs.transform === 'none' ? undefined : cs.transform);
      const tr = cs.translate === 'none' ? [0, 0] : cs.translate.split(' ').map(parseFloat);
      const ox = parseFloat(cs.left) + W / 2 + (tr[0] || 0); const oy = parseFloat(cs.top) + H / 2 + (tr[1] || 0);
      const pt = (x, y) => ({ x: ox + m.a * x + m.c * y + m.e, y: oy + m.b * x + m.d * y + m.f });
      return { nw: pt(-W / 2, -H / 2), ne: pt(W / 2, -H / 2), se: pt(W / 2, H / 2), sw: pt(-W / 2, H / 2) };
    };
    g.__hdl = () => { const { hnd } = g.__s; const R = hnd.root.getBoundingClientRect(); const o = {}; for (const h of hnd.root.querySelectorAll('.sf-sel .sf-hdl')) { const r = h.getBoundingClientRect(); o[h.dataset.h] = { x: r.left + r.width / 2 - R.left, y: r.top + r.height / 2 - R.top, w: r.width }; } return o; };
    g.__geom = () => { const { sel, hnd } = g.__s; const id = sel.id; const n = hnd.nodes.get(id); const geo = n.querySelector('.sf-geo'); return { id, c: g.__corners(n), h: g.__hdl(), size: sel.size(id), tf: n.style.transform, geoTf: geo ? geo.style.transform : null, box: [n.style.left, n.style.top, n.style.width, n.style.height].join('|'), tr: n.style.translate, resizing: sel.resizing, dragging: sel.dragging, us: hnd.root.style.userSelect, cursor: hnd.root.style.cursor }; };
    g.__worlds = () => Object.fromEntries(g.__s.sc.elements.map((e) => [e.name, { w: e.world.w / 12700, h: e.world.h / 12700, rot: e.world.rot || 0 }]));
  });
  const DIRS = { nw: [-1, -1], n: [0, -1], ne: [1, -1], e: [1, 0], se: [1, 1], s: [0, 1], sw: [-1, 1], w: [-1, 0] };
  const OPP = { nw: 'se', n: 's', ne: 'sw', e: 'w', se: 'nw', s: 'n', sw: 'ne', w: 'e' };
  const dist = (a, b) => Math.hypot(a.x - b.x, a.y - b.y);
  const fire = (type, x, y, o) => page.evaluate(({ type, x, y, o }) => globalThis.__ptr(type, x, y, o), { type, x, y, o });
  const geom = () => page.evaluate(() => globalThis.__geom());
  const worlds = () => page.evaluate(() => globalThis.__worlds());
  const fresh = async (file, name, shared = true) => { await page.evaluate(() => globalThis.__newE()); await draw(file, 0, shared); if (name) await selectByName(name); };
  const press = async (h, dx, dy, steps = 1) => { const c = (await geom()).h[h]; await fire('pointerdown', c.x, c.y); for (let i = 1; i <= steps; i++) await fire('pointermove', c.x + (dx * i) / steps, c.y + (dy * i) / steps); return c; };
  const resize = async (h, dx, dy) => { const c = await press(h, dx, dy); await fire('pointerup', c.x + dx, c.y + dy); return geom(); };
  /** What a resize by the screen delta (dx, dy) of handle h must give for a frame W, by projecting the delta on the frame's own axes (dot products, not the controller's formula). */
  const expectSize = (W, h, dx, dy) => { const t = (W.rot * Math.PI) / 180; const fx = dx * Math.cos(t) + dy * Math.sin(t); const fy = -dx * Math.sin(t) + dy * Math.cos(t); const [sx, sy] = DIRS[h]; return { w: sx && W.w >= 1 ? Math.max(2, W.w + sx * fx) : W.w, h: sy && W.h >= 1 ? Math.max(2, W.h + sy * fy) : W.h }; };
  const angDiff = (a, b) => Math.abs((((a - b) % 360) + 540) % 360 - 180);
  /** All the invariants of one finished resize in one verdict. */
  const verdict = (g0, g1, W, h, dx, dy) => {
    const e = expectSize(W, h, dx, dy); const c = g1.c;
    const topAng = (Math.atan2(c.ne.y - c.nw.y, c.ne.x - c.nw.x) * 180) / Math.PI;
    const v = {
      size: Math.abs(g1.size.w - e.w) < 1e-3 && Math.abs(g1.size.h - e.h) < 1e-3,
      anchor: dist(g1.h[OPP[h]], g0.h[OPP[h]]) < 0.05, // the opposite handle has not moved
      frame: ['nw', 'ne', 'se', 'sw'].every((k) => dist(c[k], g1.h[k]) < 0.05), // the drawn element and its handles are the same frame
      drawn: Math.abs(dist(c.nw, c.ne) - g1.size.w) < 0.05 && Math.abs(dist(c.nw, c.sw) - g1.size.h) < 0.05,
      rot: angDiff(topAng, W.rot) < 0.01, // rotation untouched
      inner: g1.geoTf === g0.geoTf && g1.box === g0.box, // the flip inside the drawing and the node's own left / top / width / height are not written
      done: g1.resizing === null && g1.dragging === null && g1.us === '',
    };
    return { all: Object.values(v).every(Boolean), v, e, got: g1.size };
  };

  // 1. every handle, unrotated element: the dragged edge(s) follow the pointer, the opposite edge / corner stays, the drawn element is the outline's frame
  for (const h of Object.keys(DIRS)) {
    await fresh(SEL, 'front-rect'); const W = (await worlds())['front-rect']; const g0 = await geom();
    const g1 = await resize(h, 37, 23); const r = verdict(g0, g1, W, h, 37, 23);
    ok(r.all, `front-rect, handle ${h}: dragged by 37 / 23 pt it gives ${r.e.w.toFixed(0)} x ${r.e.h.toFixed(0)} pt, the opposite handle stays put, element = outline = handles`, r);
  }

  // 2. rotated, flipped and group-nested elements (frame axes are the rotated ones: the pointer travel is projected on them)
  for (const name of ['g02-rot30', 'g05-flipH-rot45', 'g04-flipV', 'g07-child', 'g08-child-rot10']) {
    const res = {};
    for (const h of ['se', 'nw', 's', 'n']) {
      await fresh(GEO, name); const W = (await worlds())[name]; const g0 = await geom();
      const g1 = await resize(h, 29, -17); const r = verdict(g0, g1, W, h, 29, -17); res[h] = r.all || r;
    }
    ok(Object.values(res).every((x) => x === true), `${name}: handles se, nw, s and n follow the pointer along the element's own rotated axes; rotation, flip and group transform are preserved`, res);
  }
  {
    // the drawing itself is stretched, not re-laid out: the node carries translate(shift) <its rotation> scale(sx, sy), and the pre-existing rotation text is kept verbatim
    await fresh(GEO, 'g02-rot30'); const g0 = await geom(); const g1 = await resize('e', 40, 0);
    ok(g0.tf === 'rotate(30deg)' && /^translate\(calc\(var\(--u\) \* [-\d.]+\), calc\(var\(--u\) \* [-\d.]+\)\) rotate\(30deg\) scale\([\d.]+, 1\)$/.test(g1.tf), 'the resized node\'s transform is translate(shift) + its own rotate(30deg) + scale(sx, 1)', { before: g0.tf, after: g1.tf });
  }

  // 3. zoom: the same drag at 50 %, 100 % and 150 % gives the same size in points, the frame follows the slide, handles keep their screen size
  {
    await fresh(SEL, 'front-rect'); const zs = {};
    for (const [w, dx, dy, dw, dh] of [[480, 20, 10, 40, 20], [1440, 30, 15, 60, 30]]) {
      await page.evaluate((w) => globalThis.__s.hnd.setWidth(w), w);
      const g0 = await geom(); const g1 = await resize('se', dx, dy);
      zs[w] = { dw: g1.size.dw, dh: g1.size.dh, ok: Math.abs(g1.size.dw - dw) < 1e-3 && Math.abs(g1.size.dh - dh) < 1e-3 && ['nw', 'ne', 'se', 'sw'].every((k) => dist(g1.c[k], g1.h[k]) < 0.05) && Math.abs(g1.h.se.w - (w > 1000 ? 8 : 8)) < 0.6 && dist(g1.h.nw, g0.h.nw) < 0.05 };
    }
    ok(zs[480].ok && zs[1440].ok, 'zoom 50 % and 150 %: 20 / 10 px and 30 / 15 px drags are 40 / 20 pt and +20 / +10 pt more, anchored at the top-left, handles still 8 px', zs);
    const at150 = await geom(); await page.evaluate(() => globalThis.__s.hnd.setWidth(480)); const at50 = await geom();
    ok(['nw', 'ne', 'se', 'sw'].every((k) => dist({ x: at150.c[k].x / 3, y: at150.c[k].y / 3 }, at50.c[k]) < 0.1) && Math.abs(at50.size.w - at150.size.w) < 1e-9, 'changing the zoom after a resize rescales the resized frame exactly (the preview is kept in points)', { at150: at150.c.se, at50: at50.c.se });
  }

  // 4. limits: never inside out, never absurd; a flat axis is left alone
  {
    await fresh(SEL, 'front-rect'); let g0 = await geom(); let g1 = await resize('e', -400, 0);
    ok(Math.abs(g1.size.w - 2) < 1e-9 && g1.size.h === 200 && dist(g1.c.nw, g0.c.nw) < 0.05 && dist(g1.c.sw, g0.c.sw) < 0.05 && Object.values(g1.c).every((p) => Number.isFinite(p.x) && Number.isFinite(p.y)), 'dragging an edge past the opposite one stops at 2 pt (no negative size, no flip); the other edge stays', { size: g1.size });
    await fresh(SEL, 'front-rect'); g1 = await resize('se', -500, -500);
    ok(Math.abs(g1.size.w - 2) < 1e-9 && Math.abs(g1.size.h - 2) < 1e-9, 'a corner dragged far inside out stops at 2 x 2 pt', g1.size);
    await fresh(SEL, 'front-rect'); g1 = await resize('e', 20000, 0);
    ok(g1.size.w <= 4 * 960 + 1e-6 && g1.size.w > 960, 'a huge drag stops at four slide widths', g1.size);
    await fresh(SEL, 'h-line'); g0 = await geom(); g1 = await resize('se', 40, 30);
    ok(Math.abs(g1.size.w - 340) < 1e-3 && g1.size.h === 0 && g1.size.dh === 0 && !/NaN|Infinity/.test(g1.tf) && /scale\(1\.\d+, 1\)/.test(g1.tf) && Object.values(g1.c).every((p) => Number.isFinite(p.x)), 'a zero-height line widens but stays flat (the flat axis is not scaled: no NaN, no vertical growth)', { tf: g1.tf, size: g1.size });
    await fresh(GEO, 'g12-line-rot90'); const W = (await worlds())['g12-line-rot90']; g0 = await geom(); g1 = await resize('se', 0, 25);
    ok(Math.abs(g1.size.w - (W.w + 25)) < 1e-3 && g1.size.h === 0 && angDiff((Math.atan2(g1.c.ne.y - g1.c.nw.y, g1.c.ne.x - g1.c.nw.x) * 180) / Math.PI, 90) < 0.01, 'a flat line turned by 90 degrees grows along its own axis (down the screen) and keeps its turn', { W, size: g1.size });
  }

  // 5. cancel and release rules: the same as for a move (3.2.3-A / 3.2.3-C)
  {
    const sizeNow = async () => (await geom()).size; const prev = { w: 340, h: 220 };
    const first = async () => { await fresh(SEL, 'front-rect'); await resize('se', 40, 20); await page.evaluate(() => { globalThis.__s.resizes.length = 0; }); };
    const rs = () => page.evaluate(() => globalThis.__s.resizes.slice());
    const keep = async () => { const g = await geom(); return Math.abs(g.size.w - prev.w) < 1e-9 && Math.abs(g.size.h - prev.h) < 1e-9 && g.dragging === null && g.us === '' && g.cursor === '' && ['nw', 'ne', 'se', 'sw'].every((k) => dist(g.c[k], g.h[k]) < 0.05); };
    await first(); let g = await geom(); ok(Math.abs(g.size.w - 340) < 1e-9 && Math.abs(g.size.h - 220) < 1e-9 && JSON.stringify(await rs()) === '[]', 'a finished resize is kept (340 x 220 pt)', g.size);
    const kinds = [
      ['Escape', () => page.keyboard.press('Escape')],
      ['pointercancel', () => fire('pointercancel', 300, 300)],
      ['a second finger', () => fire('pointerdown', 100, 100, { pointerId: 2, isPrimary: false, pointerType: 'touch' })],
      ['a right button press', () => fire('pointerdown', 100, 100, { button: 2, buttons: 2 })],
      ['blur', () => page.evaluate(() => window.dispatchEvent(new Event('blur')))],
      ['selecting another element', () => page.evaluate(() => globalThis.__s.sel.select(globalThis.__s.sc.elements.find((e) => e.name === 'back-rect').id))],
    ];
    for (const [label, act] of kinds) {
      await first(); await press('e', 30, 0); g = await geom();
      const open = g.resizing === 'e' && Math.abs(g.size.w - 370) < 1e-9;
      await act(); if (label === 'selecting another element') await selectByName('front-rect');
      ok(open && await keep() && JSON.stringify(await rs()) === '["start","cancel"]', `${label} during a resize puts the frame back to the earlier kept size (340 x 220) and reports start, cancel`, { open, g: await geom(), rs: await rs() });
    }
    // a stray release of another pointer neither drops the press nor orphans the drag (3.2.3-C rule applies to resizes too)
    await first(); const c = await press('se', 20, 10); await fire('pointerup', 0, 0, { pointerId: 9, isPrimary: false }); g = await geom();
    ok(g.resizing === 'se' && Math.abs(g.size.w - 360) < 1e-9, 'a stray release of another pointer in the middle of a resize leaves it open', g);
    await fire('pointermove', c.x + 50, c.y + 40); await fire('pointerup', c.x + 50, c.y + 40); g = await geom();
    ok(g.resizing === null && Math.abs(g.size.w - 390) < 1e-9 && Math.abs(g.size.h - 260) < 1e-9 && JSON.stringify(await rs()) === '["start","end"]', 'and the real release keeps it where the pointer was (390 x 260 pt)', g.size);
    // released far from the press with no move event between: still a resize
    await first(); const c2 = (await geom()).h.se; await fire('pointerdown', c2.x, c2.y); await fire('pointerup', c2.x + 60, c2.y + 30); g = await geom();
    ok(Math.abs(g.size.w - 400) < 1e-9 && Math.abs(g.size.h - 250) < 1e-9 && JSON.stringify(await rs()) === '["start","end"]', 'a release far from the press without a move event in between is still a resize', g.size);
    // a press on a handle with no travel is a click, as before: it selects what is under the pointer and resizes nothing
    await fresh(SEL, 'front-rect'); const h = (await geom()).h.e; const under = await page.evaluate(({ x, y }) => { const { sel, sc } = globalThis.__s; const id = sel.pick(x + globalThis.__s.hnd.root.getBoundingClientRect().left, y + globalThis.__s.hnd.root.getBoundingClientRect().top); return id && sc.elements.find((e) => e.id === id).name; }, h);
    await fire('pointerdown', h.x, h.y); await fire('pointerup', h.x + 1, h.y);
    const after = await page.evaluate(() => { const { sel, sc } = globalThis.__s; return { name: sel.id && sc.elements.find((e) => e.id === sel.id).name, size: sel.id ? sel.size(sel.id) : null, resizes: globalThis.__s.resizes.slice() }; });
    ok(after.name === under && (!after.size || (after.size.dw === 0 && after.size.dh === 0)) && after.resizes.length === 0, `a click on a handle without travel is still a click: it selects what is under it (${under}) and resizes nothing`, { under, after });
  }

  // 6. preview only: the Scene Model, the XML, the edit-state and the drawn element's own box are never touched; reset / dispose / a redraw put the file back
  {
    await fresh(SEL, null); const base = await page.evaluate(async () => { const { sc, hnd, parsed } = globalThis.__s; return { rects: JSON.stringify(globalThis.__rects()), html: hnd.root.outerHTML, worlds: JSON.stringify(sc.elements.map((e) => e.world)), zip: await globalThis.__zipKey(parsed.zip), size: globalThis.__E.size, v: globalThis.__E.version }; });
    await selectByName('front-rect'); await resize('se', 40, 20); await selectByName('circle'); await resize('n', 0, -30); await selectByName('rot-rect'); await resize('e', 30, 10);
    const mid = await page.evaluate(async () => { const { sc, parsed, sel } = globalThis.__s; return { worlds: JSON.stringify(sc.elements.map((e) => e.world)), zip: await globalThis.__zipKey(parsed.zip), size: globalThis.__E.size, v: globalThis.__E.version, dirty: sc.elements.filter((e) => sel.size(e.id).dw || sel.size(e.id).dh).length }; });
    ok(mid.dirty === 3 && mid.worlds === base.worlds && mid.zip === base.zip && mid.size === base.size && mid.v === base.v, 'three resized elements: the Scene Model worlds, every part of the deck zip and the edit-state (size, version) are exactly as before', { dirty: mid.dirty, same: [mid.worlds === base.worlds, mid.zip === base.zip, mid.size === base.size, mid.v === base.v] });
    await page.evaluate(() => globalThis.__s.sel.clear());
    const cleared = await page.evaluate(() => globalThis.__s.hnd.root.outerHTML);
    ok(cleared !== base.html, 'the resize previews are drawn on the slide until reset (a cleared selection does not undo them)');
    await page.evaluate(() => globalThis.__s.sel.reset());
    const back = await page.evaluate(() => ({ html: globalThis.__s.hnd.root.outerHTML, e: globalThis.__E.size }));
    ok(back.html === base.html && back.e === 0, 'reset() puts every resized element back: the slide DOM is byte-identical to the file drawing again');
    await selectByName('front-rect'); await resize('se', 40, 20); await page.evaluate(() => globalThis.__s.sel.dispose());
    const disp = await page.evaluate(() => globalThis.__s.hnd.root.outerHTML);
    ok(disp === base.html, 'dispose() takes the resize preview off the drawn slide (DOM byte-identical, no outline, no cursor)');
    await fresh(SEL, 'front-rect'); await resize('se', 40, 20); await draw(SEL, 0, true);
    const redrawn = await page.evaluate(() => ({ rects: JSON.stringify(globalThis.__rects()), scaled: globalThis.__s.hnd.root.querySelectorAll('.sf-el[style*="scale("]').length, sel: globalThis.__s.sel.id }));
    ok(redrawn.rects === base.rects && redrawn.scaled === 0 && redrawn.sel === null, 'a redraw of the slide shows the file frames: a resize preview is not drawn again (it lives and ends with the controller)', redrawn.scaled);
  }

  // 7. a move after a resize: the element moves as drawn, its outline and handles with it; the edit-state record holds the move only
  {
    await fresh(SEL, 'front-rect'); await resize('se', 40, 20); let g = await geom();
    const cx = (g.c.nw.x + g.c.se.x) / 2; const cy = (g.c.nw.y + g.c.se.y) / 2;
    await fire('pointerdown', cx, cy); await fire('pointermove', cx + 30, cy + 12); await fire('pointermove', cx + 50, cy + 12); await fire('pointerup', cx + 50, cy + 12);
    const g2 = await geom(); const rec = await page.evaluate(() => { const { sel } = globalThis.__s; return { rec: globalThis.__E.get(sel.id), off: sel.offset(sel.id), size: globalThis.__E.size }; });
    ok(Math.abs(g2.size.w - 340) < 1e-9 && Math.abs(g2.size.h - 220) < 1e-9 && Math.abs(g2.c.nw.x - g.c.nw.x - 50) < 0.05 && Math.abs(g2.c.nw.y - g.c.nw.y - 12) < 0.05 && ['nw', 'ne', 'se', 'sw'].every((k) => dist(g2.c[k], g2.h[k]) < 0.05), 'dragging the resized element moves it by the pointer travel, keeps its size, and the handles go with it', { g: g.c.nw, g2: g2.c.nw });
    ok(rec.rec && Math.abs(rec.rec.slideDelta.dx - 50) < 1e-6 && Math.abs(rec.rec.slideDelta.dy - 12) < 1e-6 && !('w' in rec.rec) && !('h' in rec.rec) && !('size' in rec.rec) && rec.size === 1, 'the edit-state holds the 50 / 12 pt move only: no size anywhere in the record', rec.rec);
    // and a resize of an already moved element (the move offset is kept, the frame grows from the moved position)
    const g3 = await geom(); const c = g3.h.se; await fire('pointerdown', c.x, c.y); await fire('pointermove', c.x + 10, c.y + 10); await fire('pointerup', c.x + 10, c.y + 10); const g4 = await geom();
    ok(dist(g4.h.nw, g3.h.nw) < 0.05 && Math.abs(g4.size.w - 350) < 1e-9 && g4.tr === g3.tr, 'resizing a moved element keeps its move and anchors at its moved top-left corner', { tr: [g3.tr, g4.tr] });
  }

  // 8. hit-testing follows what is drawn: a stretched line is picked along its new path
  {
    await fresh(SEL, 'diag-line'); await resize('se', 100, 0);
    const r = await page.evaluate(() => { const { sel, hnd, sc } = globalThis.__s; const R = hnd.root.getBoundingClientRect(); const id = sc.elements.find((e) => e.name === 'diag-line').id; const p = (x, y) => sel.pick(R.left + x, R.top + y); return { newMid: p(550, 440) === id, oldMid: p(500, 440) === id }; });
    ok(r.newMid && !r.oldMid, 'a diagonal line stretched by 100 pt is picked at its new centre (550, 440) and no longer at its old one (500, 440)', r);
  }

  // 9. the cursor: transparent handles, so the slide (and the selected node) show the resize cursor over a handle's reach, turned with the element
  {
    await fresh(SEL, 'front-rect'); const cur = {};
    for (const h of ['e', 'se', 's', 'sw', 'nw', 'n']) { const c = (await geom()).h[h]; await fire('pointermove', c.x, c.y, { buttons: 0 }); cur[h] = await page.evaluate(() => { const { hnd, sel } = globalThis.__s; return [hnd.root.style.cursor, hnd.nodes.get(sel.id).style.cursor]; }); }
    ok(JSON.stringify(cur) === JSON.stringify({ e: ['ew-resize', 'ew-resize'], se: ['nwse-resize', 'nwse-resize'], s: ['ns-resize', 'ns-resize'], sw: ['nesw-resize', 'nesw-resize'], nw: ['nwse-resize', 'nwse-resize'], n: ['ns-resize', 'ns-resize'] }), 'hovering a handle shows its resize cursor on the slide and on the selected node (whose own cursor is `move`)', cur);
    await fire('pointermove', 300, 240, { buttons: 0 });
    const away = await page.evaluate(() => { const { hnd, sel } = globalThis.__s; return [hnd.root.style.cursor, hnd.nodes.get(sel.id).style.cursor]; });
    ok(away[0] === '' && away[1] === 'move', 'away from the handles the slide has no cursor of its own and the node shows `move` again', away);
    await fresh(GEO, 'g02-rot30'); const ce = (await geom()).h.e; await fire('pointermove', ce.x, ce.y, { buttons: 0 });
    ok((await geom()).cursor === 'nesw-resize' || (await geom()).cursor === 'nwse-resize', 'the cursor of a handle turns with the element: the east handle of a 30 degree element points along its own axis (nwse after turning 30 -> 45 steps)', (await geom()).cursor);
    await fresh(GEO, 'g12-line-rot90'); const c90 = (await geom()).h.se; await fire('pointermove', c90.x, c90.y, { buttons: 0 });
    ok(['nesw-resize', 'nwse-resize'].includes((await geom()).cursor), 'a handle on an element turned by 90 degrees gets the turned cursor', (await geom()).cursor);
    await page.evaluate(() => globalThis.__s.sel.clear());
    ok(await page.evaluate(() => globalThis.__s.hnd.root.style.cursor === '' && !globalThis.__s.hnd.root.getAttribute('style')?.includes('cursor')), 'clearing the selection leaves no cursor on the slide');
    await fresh(SEL, 'front-rect'); const ch = (await geom()).h.e; await fire('pointermove', ch.x, ch.y, { buttons: 0 }); await fire('pointermove', ch.x, ch.y, { buttons: 0, pointerType: 'touch' });
    ok((await geom()).cursor === 'ew-resize', 'a touch pointer does not change the cursor');
  }

  // 10. the handles are still transparent to input, the outline too (the contract of 3.2.2 and 3.2.1)
  {
    await fresh(SEL, 'front-rect');
    const t = await page.evaluate(() => { const { hnd } = globalThis.__s; return [...hnd.root.querySelectorAll('.sf-hdl, .sf-rot, .sf-sel, .sf-sel-h')].every((n) => getComputedStyle(n).pointerEvents === 'none'); });
    ok(t, 'all handles, the rotation knob and the outline still have pointer-events: none (a resize is found by geometry, not by the DOM)');
    const rot = await page.evaluate(() => { const { hnd } = globalThis.__s; const R = hnd.root.getBoundingClientRect(); const k = hnd.root.querySelector('.sf-rot').getBoundingClientRect(); globalThis.__ptr('pointerdown', k.left + k.width / 2 - R.left, k.top + k.height / 2 - R.top); globalThis.__ptr('pointermove', k.left + k.width / 2 - R.left + 40, k.top + k.height / 2 - R.top + 5); globalThis.__ptr('pointerup', k.left + k.width / 2 - R.left + 40, k.top + k.height / 2 - R.top + 5); const { sel } = globalThis.__s; return { resizing: sel.resizing, dragging: sel.dragging, size: sel.id ? sel.size(sel.id) : null }; });
    ok(rot.resizing === null && rot.dragging === null && (!rot.size || (rot.size.dw === 0 && rot.size.dh === 0)), 'the rotation knob is still inert: pressing and dragging it resizes nothing', rot);
  }

  // 11. code guards: the resize path writes only to the drawn nodes and outline
  {
    const code = readFileSync(path.join(ROOT, 'js/render/selection.js'), 'utf8').replace(/\/\/.*$/gm, '');
    const body = (a, b) => code.slice(code.indexOf(a), code.indexOf(b));
    const rz = body('function placeSize', 'const blockSelect') + body('function styleOutline', 'function grabbable');
    ok(!/\bedits\b/.test(rz) && !/\.world(\.\w+)?\s*[-+*/]?=[^=]/.test(rz) && !/byId\.(set|delete|clear)/.test(rz) && !/localStorage|indexedDB|fetch\(|innerHTML|cssText|setAttribute\(\s*['"]style/.test(rz), 'the resize code never reaches the edit-state, the Scene Model, storage, the network or an inline style attribute');
    ok(/const onTouchStart[^\n]*passive|addEventListener\('touchstart', onTouchStart, \{ passive: false \}\)/.test(code) && /removeEventListener\('touchstart', onTouchStart\)/.test(code), 'the touchstart listener (a finger on a handle must not scroll) is added with passive: false and removed on dispose');
    ok(/\.sf-sel-h\s*\{[^}]*pointer-events:\s*none/.test(readFileSync(path.join(ROOT, 'css/render.css'), 'utf8')), 'the handle layer is still static CSS with pointer-events: none');
  }
}

// ======================================================================================================
section('resize preview in the real UI: mouse');
{
  const ui = await ctx.newPage(); watch(ui, 'resize-ui: ');
  await ui.setViewportSize({ width: 1280, height: 900 });
  await openRendered(ui, 'Resize deck');
  const sc = (await slideBox(ui)).s;
  const info = () => ui.evaluate(() => {
    const root = document.querySelector('.rendered-host .sf-slide'); const o = root.querySelector('.sf-sel'); const node = o && [...root.querySelectorAll('.sf-el')].find((n) => n.dataset.id === o.dataset.for);
    const R = root.getBoundingClientRect(); const r = node?.getBoundingClientRect(); const q = o?.getBoundingClientRect(); const h = o && o.querySelector('.sf-hdl[data-h="se"]').getBoundingClientRect();
    return { w: r && r.width, h: r && r.height, l: r && r.left - R.left, t: r && r.top - R.top, outDx: r && Math.max(Math.abs(q.left - r.left), Math.abs(q.top - r.top), Math.abs(q.right - r.right), Math.abs(q.bottom - r.bottom)), se: h && [h.left + h.width / 2, h.top + h.height / 2], status: document.querySelector('.sf-sel-status')?.textContent || '', cursor: root.style.cursor, us: root.style.userSelect, resized: root.querySelectorAll('.sf-el[style*="scale("]').length };
  });
  await ui.mouse.click(...(await at(ui, 450, 300)));
  const a = await info();
  await ui.mouse.move(a.se[0], a.se[1]); const hov = await info();
  ok(hov.cursor === 'nwse-resize', 'real mouse: hovering the south-east handle shows the nwse resize cursor', hov.cursor);
  await ui.mouse.down(); await ui.mouse.move(a.se[0] + 30 * sc, a.se[1] + 20 * sc, { steps: 4 }); await ui.mouse.move(a.se[0] + 60 * sc, a.se[1] + 40 * sc, { steps: 4 }); const mid = await info(); await ui.mouse.up();
  const b = await info();
  ok(mid.us === 'none' && Math.abs(mid.w - a.w - 60 * sc) <= 1 && Math.abs(mid.h - a.h - 40 * sc) <= 1, 'real mouse: the element grows with the pointer while the button is down (+60 / +40 pt), text selection is off meanwhile', { a: [a.w, a.h], mid: [mid.w, mid.h], us: mid.us });
  ok(Math.abs(b.w - a.w - 60 * sc) <= 1 && Math.abs(b.h - a.h - 40 * sc) <= 1 && Math.abs(b.l - a.l) <= 0.5 && Math.abs(b.t - a.t) <= 0.5 && b.outDx <= 1.5 && b.us === '' && b.cursor === '', 'real mouse: after the release it stays 60 / 40 pt bigger, anchored at its top-left corner, the outline is the same frame', { a, b });
  ok(/Resized in the preview to 360 × 240 pt \(not saved\)/.test(b.status) && !/Moved/.test(b.status), 'the status line reports the resized size and says it is not saved', b.status);
  // a drag on the body still moves (and the two previews add up on the status line)
  const sb = await slideBox(ui); const bx = sb.x + b.l + b.w / 2; const by = sb.y + b.t + b.h / 2;
  await ui.mouse.move(bx, by); await ui.mouse.down(); await ui.mouse.move(bx + 20 * sc, by + 10 * sc, { steps: 4 }); await ui.mouse.up();
  const m = await info();
  ok(Math.abs(m.l - b.l - 20 * sc) <= 1 && Math.abs(m.w - b.w) <= 0.5 && /Moved in the preview by 20 pt, 10 pt \(not saved\)\. Resized in the preview to 360 × 240 pt/.test(m.status), 'real mouse: dragging the body of the resized element still moves it, the size is kept, the status line shows both', { m, status: m.status });
  // another slide and back: the resize preview was only on the drawn slide; the committed move comes back (3.2.3-B), the resize does not
  await ui.click('.rendered .film-item >> nth=1'); await ui.waitForSelector('.rendered-host .sf-slide[data-slide="1"] .sf-el');
  await ui.click('.rendered .film-item >> nth=0'); await ui.waitForSelector('.rendered-host .sf-slide[data-slide="0"] .sf-el');
  const back = await ui.evaluate(() => { const root = document.querySelector('.rendered-host .sf-slide'); return { resized: root.querySelectorAll('.sf-el[style*="scale("]').length, moved: [...root.querySelectorAll('.sf-el')].filter((n) => n.style.translate).length }; });
  ok(back.resized === 0 && back.moved === 1, 'another slide and back: the resize is gone (a preview ends with its slide), the committed move is drawn again', back);
  await ui.close();
}

// ======================================================================================================
section('resize preview in the real UI: touch (a finger on a handle resizes instead of scrolling)');
{
  const tctx = await browser.newContext({ viewport: { width: 1280, height: 900 }, hasTouch: true, deviceScaleFactor: 2 });
  const ui = await tctx.newPage(); watch(ui, 'resize-touch: ');
  await openRendered(ui, 'Resize touch deck');
  await ui.setViewportSize({ width: 390, height: 844 });
  await ui.waitForFunction(() => document.querySelector('.rendered-host .sf-slide').getBoundingClientRect().width < 380);
  await ui.evaluate(() => document.querySelector('.rendered-host .sf-slide').scrollIntoView({ block: 'center' }));
  const cdp = await tctx.newCDPSession(ui);
  const sc = (await slideBox(ui)).s;
  const info = () => ui.evaluate(() => {
    const root = document.querySelector('.rendered-host .sf-slide'); const o = root.querySelector('.sf-sel'); const node = o && [...root.querySelectorAll('.sf-el')].find((n) => n.dataset.id === o.dataset.for);
    const r = node?.getBoundingClientRect(); const h = o && o.querySelector('.sf-hdl[data-h="se"]').getBoundingClientRect();
    return { w: r && r.width, h: r && r.height, se: h && [h.left + h.width / 2, h.top + h.height / 2], status: document.querySelector('.sf-sel-status')?.textContent || '', scrollY: window.scrollY, hs: h && h.width };
  });
  await ui.touchscreen.tap(...(await at(ui, 450, 300)));
  const a = await info();
  ok(a.hs >= 10 - 0.01, 'phone width: the handles are 10 px (coarse pointer)', a.hs);
  const drag = async (x, y, dx, dy) => { await cdp.send('Input.dispatchTouchEvent', { type: 'touchStart', touchPoints: [{ x, y, id: 1 }] }); for (let i = 1; i <= 4; i++) await cdp.send('Input.dispatchTouchEvent', { type: 'touchMove', touchPoints: [{ x: x + (dx * i) / 4, y: y + (dy * i) / 4, id: 1 }] }); await cdp.send('Input.dispatchTouchEvent', { type: 'touchEnd', touchPoints: [] }); await ui.waitForTimeout(150); };
  await drag(a.se[0], a.se[1], 30 * sc, 20 * sc);
  const b = await info();
  ok(Math.abs(b.w - a.w - 30 * sc) <= 1 && Math.abs(b.h - a.h - 20 * sc) <= 1 && b.scrollY === a.scrollY && /Resized in the preview to 330 × 220 pt/.test(b.status), 'a finger on the south-east handle resizes the element (+30 / +20 pt) and the page does not scroll', { a, b });
  // the half of the handle that lies outside the element (and the reach around it) is still the handle: this is what the touchstart listener is for
  const c0 = await info(); await drag(c0.se[0] + 8, c0.se[1] + 8, 40 * sc, 40 * sc);
  const c = await info();
  ok(Math.abs(c.w - c0.w - 40 * sc) <= 1.5 && Math.abs(c.h - c0.h - 40 * sc) <= 1.5 && c.scrollY === a.scrollY, 'a finger that lands 8 px outside the handle, off the element, still resizes it and does not scroll the page', { c0, c });
  // far from every handle and off the element: nothing is resized (the page may scroll, as before)
  const d0 = await info(); await drag(d0.se[0] + 70, d0.se[1] + 70, 0, -20 * sc);
  const d = await info();
  ok(Math.abs(d.w - d0.w) <= 0.5 && Math.abs(d.h - d0.h) <= 0.5, 'a finger far from every handle resizes nothing', { d0, d });
  await tctx.close();
}

// ======================================================================================================
section('page health');
ok(consoleErrors.length === 0, 'no console errors / CSP violations during the run', consoleErrors.slice(0, 5));

await browser.close(); server.close();
console.log(`\n${pass} passed, ${fail} failed`);
process.exit(fail ? 1 : 0);
