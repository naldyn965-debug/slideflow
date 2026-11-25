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
    const changes = []; const moves = []; const resizes = []; const editEvents = [];
    const sel = sl.attachSelection(hnd, sc, { onChange: (id) => changes.push(id), onMove: (m) => moves.push(m.phase), onResize: (m) => resizes.push(m.phase), onEdit: (m) => editEvents.push({ ...m }), edits: shared ? globalThis.__E : undefined }); // shared: the Phase 3.2.3-B edit-state (globalThis.__E), kept across redraws
    globalThis.__s = { hnd, sc, sel, parsed, changes, moves, resizes, editEvents };
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
async function openRendered(pg, name, file = SEL) {
  await pg.goto(BASE + '/index.html'); await pg.waitForSelector('.shell');
  await pg.click('text=Templates >> nth=0'); await pg.waitForSelector('.dropzone');
  const [fc] = await Promise.all([pg.waitForEvent('filechooser'), pg.click('.dropzone')]); await fc.setFiles(file);
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
    const sel = sl.attachSelection(hnd, sc, { onMove: (m) => phases.push(m.phase), edits: { forSlide: () => [], resizesForSlide: () => [], textsForSlide: () => [], effectiveText: (e) => e.text, commitText: () => ({ ok: true, edit: null, changed: false }), discardText() {}, commit: () => { asked++; return { ok: false, reason: 'singular' }; }, commitResize: () => ({ ok: true, edit: null, changed: false }), discard() {}, discardResize() {} } }); // Phase 3.2.4-B grew the edit-state contract: the stub has the resize methods too
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
  ok(!/\.(world|xfrm|el)\s*[-+*/]?=[^=]|\bdoc\b|\.ownerDocument|byId\.(set|delete|clear)/.test(code) && [...code.matchAll(/^import .* from '(.*)';$/gm)].map((m) => m[1]).join() === './scene.js,./resize-limits.js' && !/\bimport\b|\bfetch\b|document|window/.test(readFileSync(path.join(ROOT, 'js/model/resize-limits.js'), 'utf8').replace(/\/\/.*$/gm, '')), 'edit-state.js never assigns to the Scene Model or an XML node, and imports nothing but the matrix helpers of scene.js and the constants of resize-limits.js (which imports nothing)');
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
section('resize (Phase 3.2.4-A preview, committed to the edit-state by 3.2.4-B): drag the 8 handles');
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

  // 6. the Scene Model, the XML and the drawn element's own box are never touched; reset / dispose put the file back (since 3.2.4-B a completed resize is also kept in the edit-state and a redraw draws it again)
  {
    await fresh(SEL, null); const base = await page.evaluate(async () => { const { sc, hnd, parsed } = globalThis.__s; return { rects: JSON.stringify(globalThis.__rects()), html: hnd.root.outerHTML, worlds: JSON.stringify(sc.elements.map((e) => e.world)), zip: await globalThis.__zipKey(parsed.zip), size: globalThis.__E.size, v: globalThis.__E.version }; });
    await selectByName('front-rect'); await resize('se', 40, 20); await selectByName('circle'); await resize('n', 0, -30); await selectByName('rot-rect'); await resize('e', 30, 10);
    const mid = await page.evaluate(async () => { const { sc, parsed, sel } = globalThis.__s; return { worlds: JSON.stringify(sc.elements.map((e) => e.world)), zip: await globalThis.__zipKey(parsed.zip), size: globalThis.__E.size, v: globalThis.__E.version, dirty: sc.elements.filter((e) => sel.size(e.id).dw || sel.size(e.id).dh).length }; });
    ok(mid.dirty === 3 && mid.worlds === base.worlds && mid.zip === base.zip && mid.size === base.size + 3 && mid.v === base.v + 3, 'three resized elements: the Scene Model worlds and every part of the deck zip are exactly as before; since 3.2.4-B the edit-state holds the three completed resizes (3 elements, version +3) and nothing else', { dirty: mid.dirty, same: [mid.worlds === base.worlds, mid.zip === base.zip, mid.size === base.size + 3, mid.v === base.v + 3] });
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
    ok(redrawn.rects !== base.rects && redrawn.scaled === 1 && redrawn.sel === null, 'a redraw of the slide draws the COMMITTED resize again (3.2.4-B; before it showed the file frame): one stretched element, nothing selected', redrawn.scaled);
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
  // another slide and back: the committed move (3.2.3-B) and the committed resize (3.2.4-B) come back
  await ui.click('.rendered .film-item >> nth=1'); await ui.waitForSelector('.rendered-host .sf-slide[data-slide="1"] .sf-el');
  await ui.click('.rendered .film-item >> nth=0'); await ui.waitForSelector('.rendered-host .sf-slide[data-slide="0"] .sf-el');
  const back = await ui.evaluate(() => { const root = document.querySelector('.rendered-host .sf-slide'); return { resized: root.querySelectorAll('.sf-el[style*="scale("]').length, moved: [...root.querySelectorAll('.sf-el')].filter((n) => n.style.translate).length }; });
  ok(back.resized === 1 && back.moved === 1, 'another slide and back: the committed resize (3.2.4-B) and the committed move (3.2.3-B) are both drawn again', back);
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
section('resize commit (Phase 3.2.4-B): the edit-state keeps completed resizes, in the parent group\'s coordinates');
await page.evaluate(() => {
  /** Select `name`, press the centre of its handle `h` and drag by (dx, dy) slide px (pt at this zoom); returns the size afterwards, or null when the element / handle is not there (small-frame rule). */
  globalThis.__resizeBy = (name, h, dx, dy, release = true) => {
    const { sel, sc, hnd } = globalThis.__s; const e = sc.elements.find((x) => x.name === name); if (!e || !sel.select(e.id)) return null;
    const R = hnd.root.getBoundingClientRect(); const hd = hnd.root.querySelector(`.sf-sel .sf-hdl[data-h="${h}"]`); const r = hd && hd.getBoundingClientRect(); if (!r || !(r.width > 0)) return null;
    const x = r.left + r.width / 2 - R.left; const y = r.top + r.height / 2 - R.top;
    globalThis.__ptr('pointerdown', x, y); globalThis.__ptr('pointermove', x + dx / 2, y + dy / 2); globalThis.__ptr('pointermove', x + dx, y + dy);
    if (release) globalThis.__ptr('pointerup', x + dx, y + dy);
    return { id: e.id, ...sel.size(e.id), x, y };
  };
  /** Drag the body of `name` (the painted spot nearest its centre) by (dx, dy). */
  globalThis.__moveBy = (name, dx, dy, release = true) => {
    const { sel, sc, hnd } = globalThis.__s; const e = sc.elements.find((x) => x.name === name); if (!e || !sel.select(e.id)) return false;
    const R = hnd.root.getBoundingClientRect(); const r = hnd.nodes.get(e.id).getBoundingClientRect(); let best = null;
    for (let i = 1; i < 20; i++) for (let j = 1; j < 20; j++) { const x = r.left + (r.width * i) / 20; const y = r.top + (r.height * j) / 20; if (sel.pick(x, y) === e.id) { const d = Math.hypot(x - (r.left + r.width / 2), y - (r.top + r.height / 2)); if (!best || d < best.d) best = { x: x - R.left, y: y - R.top, d }; } }
    if (!best) return false;
    globalThis.__ptr('pointerdown', best.x, best.y); globalThis.__ptr('pointermove', best.x + dx / 2, best.y + dy / 2); globalThis.__ptr('pointermove', best.x + dx, best.y + dy);
    if (release) globalThis.__ptr('pointerup', best.x + dx, best.y + dy);
    return true;
  };
  /**
   * The independent check of the committed frames: copy the deck IN MEMORY, write every committed element's frame (edit-state frame(): own <a:off> and, when resized, <a:ext>, rounded to whole EMU like a file
   * holds them) into that copy, build the Scene Model of the copy with the normal forward code, and compare every element's world frame (centre, size, rotation, mirroring) with what the drag drew: the file
   * frame + the committed move + the committed resize. The original file is only read.
   */
  globalThis.__verifyFrames = async (b64, idx) => {
    const { parsed, scene, parser } = await globalThis.__load(b64); const E = 12700;
    const sc0 = await scene.buildScene(parsed, idx); const es = globalThis.__E;
    const ids = [...new Set([...es.list().map((r) => r.id), ...es.listResizes().map((r) => r.id)])].filter((id) => es.frame(id)?.part === sc0.part);
    const zip = await JSZip.loadAsync(Uint8Array.from(atob(b64), (c) => c.charCodeAt(0)));
    const doc = new DOMParser().parseFromString(await zip.file(sc0.part).async('string'), 'application/xml');
    const byXml = new Map([...doc.getElementsByTagNameNS('*', 'cNvPr')].map((c) => [c.getAttribute('id'), c]));
    for (const id of ids) {
      const f = es.frame(id); const e = sc0.byId.get(id); const xf = byXml.get(e.xmlId).parentNode.parentNode.getElementsByTagNameNS('*', 'xfrm')[0];
      const off = xf.getElementsByTagNameNS('*', 'off')[0]; const ext = xf.getElementsByTagNameNS('*', 'ext')[0];
      off.setAttribute('x', String(Math.round(f.x))); off.setAttribute('y', String(Math.round(f.y)));
      if (f.w != null) { ext.setAttribute('cx', String(Math.round(f.w))); ext.setAttribute('cy', String(Math.round(f.h))); }
    }
    zip.file(sc0.part, new XMLSerializer().serializeToString(doc));
    const sc1 = await scene.buildScene(await parser.parsePptx(await zip.generateAsync({ type: 'uint8array' })), idx);
    const angDiff = (a, b) => Math.abs((((a - b) % 360) + 540) % 360 - 180);
    const rows = [];
    for (const e of sc0.elements) {
      const e1 = sc1.byId.get(e.id); const rz = es.getResize(e.id); const mv = es.get(e.id); const W = e.world; const W1 = e1 && e1.world;
      if (!W || !W1) { rows.push({ name: e.name, err: Infinity }); continue; }
      const cx = W.cx + (rz ? rz.slideFrame.tx * E : 0) + (mv ? mv.slideDelta.dx * E : 0); const cy = W.cy + (rz ? rz.slideFrame.ty * E : 0) + (mv ? mv.slideDelta.dy * E : 0);
      const w = rz ? rz.slideFrame.w * E : W.w; const h = rz ? rz.slideFrame.h * E : W.h;
      rows.push({ name: e.name, kind: e.kind, depth: e.depth, moved: !!mv, resized: !!rz, shear: !!W.shear, rot: W.rot || 0, flipX: W.flipX, err: Math.max(Math.abs(W1.cx - cx), Math.abs(W1.cy - cy), Math.abs(W1.w - w), Math.abs(W1.h - h)), rotErr: angDiff(W1.rot || 0, W.rot || 0), flipSame: W1.flipX === W.flipX, group: e.kind === 'group' });
    }
    return { rows, edited: ids.length };
  };
});
const fireP = (type, x, y, o) => page.evaluate(({ type, x, y, o }) => globalThis.__ptr(type, x, y, o), { type, x, y, o });
const closeTo = (a, b, tol = 0.5) => Math.abs(a - b) <= tol;

// ---- 1. the pure model: records, replacement, removal, refusals, the composite frame ------------------------------------------------------------
{
  const r = await page.evaluate(async (b) => {
    const { parsed, scene } = await globalThis.__load(b); const m = await import('/js/model/edit-state.js'); const sc = await scene.buildScene(parsed, 0); const E = 12700;
    const el = (n) => sc.elements.find((e) => e.name === n); const id = (n) => el(n).id; const fw = (n) => el(n).world.w / E; const fh = (n) => el(n).world.h / E;
    const es = m.createEditState(); const out = {};
    // a top-level element, the shape a south-east drag gives: +30 / +10 pt, the centre moves by half of that
    const a = es.commitResize(sc, id('g01-plain'), { w: fw('g01-plain') + 30, h: fh('g01-plain') + 10, tx: 15, ty: 5 });
    out.a = { ok: a.ok, changed: a.changed, v: es.version, keys: Object.keys(a.edit).sort().join(), edit: a.edit, plain: JSON.parse(JSON.stringify(a.edit)) };
    // the next commit is the TOTAL since the file (a north-west drag this time: the centre moves the other way)
    const b2 = es.commitResize(sc, id('g01-plain'), { w: fw('g01-plain') + 30, h: fh('g01-plain') + 10, tx: -15, ty: -5 });
    out.b = { ok: b2.ok, changed: b2.changed, v: es.version, size: es.size, pos: b2.edit.position, ext: b2.edit.extent, shift: b2.edit.localShift };
    const same = es.commitResize(sc, id('g01-plain'), { w: fw('g01-plain') + 30, h: fh('g01-plain') + 10, tx: -15, ty: -5 }); out.same = { ok: same.ok, changed: same.changed, v: es.version };
    // removal: null, and a frame that equals the file's
    const z = es.commitResize(sc, id('g01-plain'), null); out.zero = { ok: z.ok, edit: z.edit, changed: z.changed, size: es.size, get: es.getResize(id('g01-plain')), v: es.version };
    es.commitResize(sc, id('g01-plain'), { w: fw('g01-plain') + 30, h: fh('g01-plain'), tx: 15, ty: 0 }); const vBefore = es.version;
    const back = es.commitResize(sc, id('g01-plain'), { w: fw('g01-plain'), h: fh('g01-plain'), tx: 0, ty: 0 }); out.back = { ok: back.ok, edit: back.edit, changed: back.changed, size: es.size, v: es.version, vBefore };
    const none = es.commitResize(sc, id('g01-plain'), null); out.noneAgain = { ok: none.ok, changed: none.changed, v: es.version };
    // refusals write nothing
    const g1 = { w: fw('g01-plain') + 5, h: fh('g01-plain'), tx: 2, ty: 0 }; const v0 = es.version; const refuse = {};
    refuse.group = es.commitResize(sc, id('g06-group-scaled'), { w: 50, h: 50, tx: 0, ty: 0 }).reason; refuse.unknown = es.commitResize(sc, 'nope', g1).reason;
    refuse.nan = es.commitResize(sc, id('g01-plain'), { ...g1, w: NaN }).reason; refuse.inf = es.commitResize(sc, id('g01-plain'), { ...g1, tx: Infinity }).reason; refuse.neg = es.commitResize(sc, id('g01-plain'), { ...g1, h: -3 }).reason;
    refuse.missing = es.commitResize(sc, id('g01-plain'), { w: 50, h: 50 }).reason;
    refuse.tooSmall = es.commitResize(sc, id('g01-plain'), { ...g1, w: 1 }).reason; refuse.tooSmall2 = es.commitResize(sc, id('g01-plain'), { ...g1, h: 1.999 }).reason;
    refuse.tooLarge = es.commitResize(sc, id('g01-plain'), { ...g1, w: 4 * 960 + 0.01 }).reason; refuse.tooLargeH = es.commitResize(sc, id('g01-plain'), { ...g1, h: 4 * 540 + 0.01 }).reason;
    refuse.flat = es.commitResize(sc, id('g12-line-rot90'), { w: fw('g12-line-rot90') + 10, h: 5, tx: 5, ty: 0 }).reason;
    const e6 = sc.byId.get(id('g06-child1')); const fake = { index: 0, width: sc.width, height: sc.height, byId: new Map([['x', { ...e6, id: 'x', world: { ...e6.world, matrix: [0, 0, 0, 0, 5, 5] } }]]) };
    refuse.singular = es.commitResize(fake, 'x', { w: fw('g06-child1') + 4, h: fh('g06-child1'), tx: 2, ty: 0 }).reason;
    refuse.unplaced = es.commitResize({ index: 0, width: sc.width, height: sc.height, byId: new Map([['u', { ...e6, id: 'u', world: null }]]) }, 'u', g1).reason;
    out.refuse = refuse; out.afterRefusals = { size: es.size, v: es.version, v0, has: es.hasResize('x') };
    // the limits themselves are accepted: exactly 2 pt, exactly four slide widths / heights, an axis already under 2 pt keeps its file size
    const lim = {};
    lim.min = es.commitResize(sc, id('g01-plain'), { w: 2, h: 2, tx: 0, ty: 0 }).ok; lim.minEdge = es.getResize(id('g01-plain')).slideFrame;
    lim.maxW = es.commitResize(sc, id('g01-plain'), { w: 4 * 960, h: fh('g01-plain'), tx: 0, ty: 0 }).ok; lim.maxH = es.commitResize(sc, id('g01-plain'), { w: fw('g01-plain'), h: 4 * 540, tx: 0, ty: 0 }).ok;
    const thin = { index: 0, width: sc.width, height: sc.height, byId: new Map([['t', { ...e6, id: 't', world: { ...e6.world, w: 1.5 * E }, xfrm: { ...e6.xfrm } }]]) };
    lim.thinKeep = es.commitResize(thin, 't', { w: 1.5, h: fh('g06-child1'), tx: 0, ty: 3 }).ok; lim.thinBigger = es.commitResize(thin, 't', { w: 1.5000001 + 10, h: fh('g06-child1'), tx: 0, ty: 3 }).ok; lim.thinSmaller = es.commitResize(thin, 't', { w: 1.2, h: fh('g06-child1'), tx: 0, ty: 3 }).reason;
    const flatOk = es.commitResize(sc, id('g12-line-rot90'), { w: fw('g12-line-rot90') + 10, h: 0, tx: 5, ty: 0 }); lim.flat = { ok: flatOk.ok, sy: flatOk.edit && flatOk.edit.scale.sy, extH: flatOk.edit && flatOk.edit.extent.h, originH: flatOk.edit && flatOk.edit.origin.h, extW: flatOk.edit && flatOk.edit.extent.w, originW: flatOk.edit && flatOk.edit.origin.w };
    out.lim = lim; es.clear();
    // copies, listings, the composite frame, discard
    es.commitResize(sc, id('g01-plain'), { w: fw('g01-plain') + 30, h: fh('g01-plain') + 10, tx: 15, ty: 5 });
    const c = es.getResize(id('g01-plain')); c.slideFrame.w = 999; c.origin.x = 1; c.localShift.dx = 1; out.isolated = es.getResize(id('g01-plain')).slideFrame.w === fw('g01-plain') + 30 && es.getResize(id('g01-plain')).origin.x === 500000;
    out.only = { frame: es.frame(id('g01-plain')), moves: es.list().length, resizes: es.listResizes().length, forSlide: es.forSlide(sc.part).length, rzForSlide: es.resizesForSlide(sc.part).length, other: es.resizesForSlide('ppt/slides/slide9.xml').length, has: es.has(id('g01-plain')), hasRz: es.hasResize(id('g01-plain')), get: es.get(id('g01-plain')) };
    es.commit(sc, id('g01-plain'), 10, 5); es.commit(sc, id('g02-rot30'), 7, 7); es.commitResize(sc, id('g03-flipH'), { w: fw('g03-flipH') + 12, h: fh('g03-flipH'), tx: 6, ty: 0 });
    out.both = { frame: es.frame(id('g01-plain')), moveOnly: es.frame(id('g02-rot30')), rzOnly: es.frame(id('g03-flipH')), none: es.frame(id('g04-flipV')), size: es.size, moves: es.list().length, resizes: es.listResizes().length, v: es.version };
    es.discardResize(id('g01-plain')); out.afterDiscardRz = { rz: es.hasResize(id('g01-plain')), mv: es.has(id('g01-plain')) && !!es.get(id('g01-plain')), frame: es.frame(id('g01-plain')), size: es.size };
    es.commitResize(sc, id('g01-plain'), { w: fw('g01-plain') + 30, h: fh('g01-plain') + 10, tx: 15, ty: 5 }); es.discard(id('g01-plain')); out.afterDiscard = { has: es.has(id('g01-plain')), size: es.size };
    es.clear(); out.cleared = { size: es.size, resizes: es.listResizes().length, moves: es.list().length };
    out.deck = { same: m.editStateOf(parsed) === m.editStateOf(parsed), other: m.editStateOf(parsed) !== m.editStateOf({}) };
    return out;
  }, b64(GEO));
  const EM = 12700; const o1 = { x: 500000, y: 400000, w: 1500000, h: 800000 };
  ok(r.a.ok && r.a.changed && r.a.v === 1 && r.a.keys === ['extent', 'id', 'kind', 'localShift', 'origin', 'parentId', 'part', 'position', 'posSource', 'scale', 'slideFrame', 'slideIndex'].sort().join() && JSON.stringify(r.a.plain) === JSON.stringify(r.a.edit), 'commitResize keeps one plain-data record per element (id, part, slide index, parent id, slide frame, scale, origin, extent, local shift, position), version 1', r.a.keys);
  ok(JSON.stringify(r.a.edit.origin) === JSON.stringify(o1) && closeTo(r.a.edit.extent.w, o1.w + 30 * EM, 1e-6) && closeTo(r.a.edit.extent.h, o1.h + 10 * EM, 1e-6) && closeTo(r.a.edit.scale.sx, (o1.w / EM + 30) / (o1.w / EM), 1e-12) && closeTo(r.a.edit.localShift.dx, 15 * EM, 1e-6) && closeTo(r.a.edit.localShift.dy, 5 * EM, 1e-6), 'a top-level element: the record carries the file\'s own <a:off> / <a:ext>, the new extent is the file extent + 30 / + 10 pt (hand-computed), the local shift IS the slide shift', r.a.edit);
  ok(closeTo(r.a.edit.position.x, o1.x, 1e-6) && closeTo(r.a.edit.position.y, o1.y, 1e-6) && r.a.edit.parentId === null, 'a south-east resize (the centre moves by half the growth) keeps the file\'s top-left <a:off>: position = origin', r.a.edit.position);
  ok(r.b.ok && r.b.changed && r.b.v === 2 && r.b.size === 1 && closeTo(r.b.pos.x, o1.x - 30 * EM, 1e-6) && closeTo(r.b.pos.y, o1.y - 10 * EM, 1e-6) && closeTo(r.b.ext.w, o1.w + 30 * EM, 1e-6), 'a second commit replaces the first, it is the TOTAL since the file: a north-west resize moves the top-left <a:off> by -30 / -10 pt (hand-computed), still one record, version 2', r.b);
  ok(r.same.ok && r.same.changed === false && r.same.v === 2, 'committing the same frame again changes nothing (version unchanged)', r.same);
  ok(r.zero.ok && r.zero.edit === null && r.zero.changed === true && r.zero.size === 0 && r.zero.get === null && r.zero.v === 3, 'committing null removes the record: the element is back at its file size', r.zero);
  ok(r.back.ok && r.back.edit === null && r.back.changed === true && r.back.size === 0 && r.back.v === r.back.vBefore + 1 && r.noneAgain.ok && r.noneAgain.changed === false && r.noneAgain.v === r.back.v, 'committing a frame equal to the file\'s removes the record as well; removing what is not there changes nothing', { back: r.back, noneAgain: r.noneAgain });
  const want = { group: 'not-resizable', unknown: 'not-resizable', nan: 'invalid', inf: 'invalid', neg: 'invalid', missing: 'invalid', tooSmall: 'too-small', tooSmall2: 'too-small', tooLarge: 'too-large', tooLargeH: 'too-large', flat: 'flat-axis', singular: 'singular', unplaced: 'unplaced' };
  ok(JSON.stringify(r.refuse) === JSON.stringify(want), 'refusals, each with its reason: a group / unknown id, NaN / Infinity / negative / missing numbers, under 2 pt, over four slide sizes, a flat axis changed, a parent that cannot be inverted, an element the scene did not place', r.refuse);
  ok(r.afterRefusals.size === 0 && r.afterRefusals.v === r.afterRefusals.v0 && r.afterRefusals.has === false, 'a refusal writes nothing (size and version unchanged)', r.afterRefusals);
  ok(r.lim.min && r.lim.minEdge.w === 2 && r.lim.minEdge.h === 2 && r.lim.maxW && r.lim.maxH, 'the limits themselves are accepted: exactly 2 x 2 pt, exactly four slide widths, exactly four slide heights', r.lim);
  ok(r.lim.thinKeep && r.lim.thinBigger && r.lim.thinSmaller === 'too-small', 'an axis the file already has under 2 pt (but not flat) may keep its size or grow, never get smaller', r.lim);
  ok(r.lim.flat.ok && r.lim.flat.sy === 1 && r.lim.flat.extH === 0 && r.lim.flat.originH === 0 && closeTo(r.lim.flat.extW, r.lim.flat.originW + 10 * EM, 1e-6), 'a flat line (height 0): a resize that widens it is kept, its height scale is exactly 1 and the extent height stays 0', r.lim.flat);
  ok(r.isolated, 'records are handed out as copies: changing one does not change the edit-state');
  ok(r.only.frame && r.only.frame.moved === false && r.only.frame.resized === true && closeTo(r.only.frame.x, o1.x, 1e-6) && closeTo(r.only.frame.w, o1.w + 30 * EM, 1e-6) && r.only.moves === 0 && r.only.resizes === 1 && r.only.forSlide === 0 && r.only.rzForSlide === 1 && r.only.other === 0 && r.only.get === null && r.only.has && r.only.hasRz, 'a resize is listed with the resizes only: list() / forSlide() / get() are still the moves of 3.2.3-B (none here), listResizes() / resizesForSlide() / getResize() the resizes, has() is true for either', r.only);
  ok(r.both.size === 3 && r.both.moves === 2 && r.both.resizes === 2 && r.both.v > 0, 'size counts the elements that hold a move or a resize (3: one with both, one move, one resize)', r.both);
  ok(r.both.frame.moved && r.both.frame.resized && closeTo(r.both.frame.x, o1.x + 10 * EM, 1e-6) && closeTo(r.both.frame.y, o1.y + 5 * EM, 1e-6) && closeTo(r.both.frame.w, o1.w + 30 * EM, 1e-6), 'frame() puts a move and a resize of one element together: the move (10 / 5 pt) and the resize add up (hand-computed <a:off> and <a:ext>)', r.both.frame);
  ok(r.both.moveOnly.moved && !r.both.moveOnly.resized && r.both.moveOnly.w === null && r.both.rzOnly.resized && !r.both.rzOnly.moved && r.both.none === null, 'frame() of a move-only element has no extent (the file\'s stays), of a resize-only element no move, of an untouched element it is null', r.both);
  ok(r.afterDiscardRz.rz === false && r.afterDiscardRz.mv === true && r.afterDiscardRz.frame && r.afterDiscardRz.frame.resized === false && r.afterDiscard.has === false && r.cleared.size === 0 && r.cleared.resizes === 0 && r.cleared.moves === 0, 'discardResize() forgets the resize only (the move stays), discard() forgets move and resize, clear() forgets everything', { r1: r.afterDiscardRz, r2: r.afterDiscard, r3: r.cleared });
  ok(r.deck.same && r.deck.other, 'one edit-state per opened deck, resizes included (it is the same object)');
}

// ---- 2. group-local conversion against hand-computed values (the same group maths as a move: the group\'s rotation, flips and scale are undone) ----
{
  const r = await page.evaluate(async (b) => {
    const { parsed, scene } = await globalThis.__load(b); const m = await import('/js/model/edit-state.js'); const sc = await scene.buildScene(parsed, 0); const E = 12700;
    const es = m.createEditState(); const out = {};
    for (const [n, dW, dH, tx, ty] of [['g06-child1', 20, 8, 10, 4], ['g06-child2-rot20', 12, 6, -5, 9], ['g07-child', 20, 8, 10, 4], ['g09-child', 20, 8, 10, 4], ['g08-child-rot10', 20, 8, 10, 4], ['g02-rot30', 20, 8, 10, 4], ['g05-flipH-rot45', 20, 8, 10, 4]]) {
      const e = sc.elements.find((x) => x.name === n); const res = es.commitResize(sc, e.id, { w: e.world.w / E + dW, h: e.world.h / E + dH, tx, ty });
      out[n] = { ok: res.ok, edit: res.edit, depth: e.depth, parent: e.parentId, fw: e.world.w / E, fh: e.world.h / E, shear: !!e.world.shear };
    }
    return out;
  }, b64(GEO));
  const rad = (d) => (d * Math.PI) / 180; const R = (deg, x, y) => [Math.cos(rad(deg)) * x - Math.sin(rad(deg)) * y, Math.sin(rad(deg)) * x + Math.cos(rad(deg)) * y]; const EM = 12700;
  const g6 = r['g06-child1']; const g6b = r['g06-child2-rot20'];
  ok(g6.ok && closeTo(g6.edit.localShift.dx, 2 * 10 * EM, 1e-3) && closeTo(g6.edit.localShift.dy, 2 * 4 * EM, 1e-3) && closeTo(g6.edit.extent.w, g6.edit.origin.w + 2 * 20 * EM, 1e-3) && closeTo(g6.edit.extent.h, g6.edit.origin.h + 2 * 8 * EM, 1e-3) && closeTo(g6.edit.origin.w, 2 * g6.fw * EM, 1e-3), 'child of a group drawn at scale 0.5: the extent grows by twice the slide growth (+40 / +16 pt) and the centre shift is twice the slide shift (hand-computed)', g6.edit);
  ok(g6b.ok && closeTo(g6b.edit.localShift.dx, 2 * -5 * EM, 1e-3) && closeTo(g6b.edit.localShift.dy, 2 * 9 * EM, 1e-3) && closeTo(g6b.edit.extent.w, g6b.edit.origin.w + 2 * 12 * EM, 1e-3), 'the same for a child that is itself rotated 20 degrees (its own rotation does not enter)', g6b.edit);
  { const [ex, ey] = R(-30, 10 * EM, 4 * EM); const g = r['g07-child']; ok(g.ok && closeTo(g.edit.localShift.dx, ex, 1e-3) && closeTo(g.edit.localShift.dy, ey, 1e-3) && closeTo(g.edit.extent.w, g.edit.origin.w + 20 * EM, 1e-3) && closeTo(g.edit.extent.h, g.edit.origin.h + 8 * EM, 1e-3), 'child of a group rotated 30 degrees: the centre shift is the slide shift turned back by 30 degrees, the extent grows by the slide growth (hand-computed)', g.edit); }
  { const g = r['g09-child']; ok(g.ok && closeTo(g.edit.localShift.dx, -10 * EM, 1e-3) && closeTo(g.edit.localShift.dy, 4 * EM, 1e-3) && closeTo(g.edit.extent.w, g.edit.origin.w + 20 * EM, 1e-3), 'child of a horizontally flipped group: the horizontal part of the shift is mirrored, the extent still grows by the slide growth (a flip keeps lengths)', g.edit); }
  { const [ax, ay] = R(-20, 10 * EM, 4 * EM); const [ex, ey] = R(-40, -ax, ay); const g = r['g08-child-rot10']; ok(g.ok && g.depth >= 2 && closeTo(g.edit.localShift.dx, ex, 1e-3) && closeTo(g.edit.localShift.dy, ey, 1e-3) && closeTo(g.edit.extent.w, g.edit.origin.w + 20 * EM, 1e-3) && closeTo(g.edit.extent.h, g.edit.origin.h + 8 * EM, 1e-3), 'child nested in two groups (scale, rotation, flip, scale, rotation): shift = R(-40) F R(-20) of the slide shift, the extent follows the net scale 1 (hand-computed)', { edit: g.edit, depth: g.depth }); }
  ok(r['g02-rot30'].ok && closeTo(r['g02-rot30'].edit.localShift.dx, 10 * EM, 1e-3) && closeTo(r['g02-rot30'].edit.extent.w, r['g02-rot30'].edit.origin.w + 20 * EM, 1e-3) && r['g05-flipH-rot45'].ok && closeTo(r['g05-flipH-rot45'].edit.localShift.dy, 4 * EM, 1e-3) && r['g05-flipH-rot45'].edit.parentId === null, 'a top-level element that is itself rotated / flipped: its own rotation and flips do not change the conversion (shift = slide shift, extent = file extent + slide growth)');
}

// ---- 3. real mouse: every element of p31-geometry resized with a handle, committed, and the frames checked against a Scene Model rebuilt from a copy of the deck with the committed <a:off> / <a:ext> written into it ----
let rzDone = [];
{
  await page.evaluate(() => globalThis.__newE());
  await draw(GEO, 0, true);
  const html0 = await page.evaluate(() => globalThis.__s.hnd.root.outerHTML);
  const names = await page.evaluate(() => globalThis.__s.sc.elements.filter((e) => e.kind !== 'group' && e.kind !== 'alt').map((e) => e.name));
  const HS = ['se', 'nw', 's', 'e', 'n', 'sw', 'ne', 'w']; const skipped = []; const noop = [];
  for (const [i, name] of names.entries()) {
    const r = await page.evaluate(({ name, h }) => globalThis.__resizeBy(name, h, 29, 17), { name, h: HS[i % HS.length] });
    if (!r) { skipped.push(name); continue; }
    if (!r.dw && !r.dh) { noop.push(name); continue; } // a handle that cannot change this element (a flat line's vertical handle): nothing to commit
    rzDone.push({ name, h: HS[i % HS.length], dw: r.dw, dh: r.dh });
  }
  const recs = await page.evaluate(() => globalThis.__E.listResizes());
  ok(rzDone.length >= 11 && recs.length === rzDone.length, `each completed resize was committed: ${recs.length} elements of p31-geometry (plain, rotated, flipped, group children, nested groups, lines, text) hold a resize in the edit-state, one record per element`, { done: rzDone.length, recs: recs.length, skipped, noop });
  ok(recs.every((r) => JSON.parse(JSON.stringify(r)) && Object.values(r).every((v) => typeof v !== 'function')) && (await page.evaluate(() => globalThis.__E.list().length)) === 0, 'the records are plain data, and the move list is still empty (a resize is not a move)');
  ok(recs.every((r) => r.slideFrame.w >= 2 - 1e-9 && r.slideFrame.h >= 0), 'every kept frame respects the limits (no frame under 2 pt on a scaled axis)', recs.filter((r) => r.slideFrame.w < 2).map((r) => r.id));
  const v = await page.evaluate((b) => globalThis.__verifyFrames(b, 0), b64(GEO));
  const rows = v.rows; const rz = rows.filter((r) => r.resized);
  ok(v.edited === recs.length && rz.length === recs.length, `${v.edited} committed elements were written into an in-memory copy of the deck (own <a:off> and <a:ext>, whole EMU) and the Scene Model was rebuilt from it`, { edited: v.edited, rz: rz.length });
  ok(rz.every((r) => r.err < 16 && r.rotErr < 1e-6 && r.flipSame), 'every resized element\'s rebuilt world frame (centre, size, rotation, mirroring) equals the frame the drag drew, within rounding to whole EMU (< 16 EMU = 0.001 pt)', rz.filter((r) => !(r.err < 16 && r.rotErr < 1e-6 && r.flipSame)).map((r) => [r.name, r.err, r.rotErr]));
  ok(rows.filter((r) => !r.resized).every((r) => r.err < 16 && r.rotErr < 1e-6), 'and everything that was not resized (the groups themselves included) is exactly where it was', rows.filter((r) => !r.resized && !(r.err < 16)).map((r) => [r.name, r.err]));
  ok(rz.some((r) => r.depth >= 2) && rz.some((r) => r.depth === 1) && rz.some((r) => r.rot && r.depth === 0) && rz.some((r) => r.flipX) && rz.some((r) => /^g1[01]-/.test(r.name)) && rz.some((r) => r.name === 'g13-textbox-rot350'), 'the set contains top-level, rotated, flipped, group-child, two-groups-deep, line and text elements', rz.map((r) => [r.name, r.depth, r.rot, r.flipX]));
  // handles other than the south-east one, and the opposite edge / corner stays where it was: checked for the elements the previous loop resized with n / nw / sw / w / ne (sizes against the file are above, positions via the rebuilt scene)
  // a move on top of a resize: the composite frame still equals the drawn one
  const movedNames = rz.filter((r) => r.kind !== 'line').slice(0, 6).map((r) => r.name); let movedOk = 0;
  for (const n of movedNames) if (await page.evaluate((n) => globalThis.__moveBy(n, 23, 11), n)) movedOk++;
  const v2 = await page.evaluate((b) => globalThis.__verifyFrames(b, 0), b64(GEO));
  const both = v2.rows.filter((r) => r.moved && r.resized);
  ok(movedOk >= 4 && both.length === movedOk && v2.rows.filter((r) => r.moved || r.resized).every((r) => r.err < 16 && r.rotErr < 1e-6 && r.flipSame), `${both.length} elements moved after they were resized: the composite frame (move + resize) written into a copy of the deck still rebuilds exactly the frame that is drawn`, v2.rows.filter((r) => (r.moved || r.resized) && !(r.err < 16)).map((r) => [r.name, r.err]));
  // the slide shows the resized frames, and dispose takes them off again while the edit-state keeps every one
  const drawn = await page.evaluate(() => globalThis.__s.hnd.root.querySelectorAll('.sf-el[style*="scale("]').length);
  await page.evaluate(() => globalThis.__s.sel.dispose());
  const after = await page.evaluate(() => ({ html: globalThis.__s.hnd.root.outerHTML, n: globalThis.__E.listResizes().length }));
  ok(drawn === rzDone.length && after.html === html0, `the slide shows ${drawn} stretched elements, and dispose() takes every resize and move preview off the drawn slide (byte-identical to the file drawing) ...`, { drawn, same: after.html === html0 });
  ok(after.n === recs.length, '... while the edit-state still holds every committed resize', after.n);
}

// ---- 4. persistence: redraw, other slide and back, a total that accumulates, resizing back to the file size ----------------------------------------
{
  const rects = () => page.evaluate(() => globalThis.__rects());
  const sameRects = (a, b) => Object.keys(a).length === Object.keys(b).length && Object.keys(a).every((k) => b[k] && ['l', 't', 'w', 'h'].every((p) => Math.abs(a[k][p] - b[k][p]) < 0.01) && a[k].tf === b[k].tf && a[k].tr === b[k].tr);
  // (a) a brand-new parse, scene and drawing of the same slide with the edit-state of section 3
  await draw(GEO, 0, true); const re = await page.evaluate(() => ({ rects: globalThis.__rects(), sel: globalThis.__s.sel.id, outlines: globalThis.__s.hnd.root.querySelectorAll('.sf-sel').length, rz: globalThis.__E.listResizes().length, mv: globalThis.__E.list().length, v: globalThis.__E.version, scaled: globalThis.__s.hnd.root.querySelectorAll('.sf-el[style*="scale("]').length }));
  ok(re.scaled === rzDone.length && re.sel === null && re.outlines === 0, `a redrawn slide draws the committed resizes before anything is selected (${re.scaled} stretched elements), with no selection`, re);
  // the same drawing, drawn a second time from yet another parse: identical to the first redraw (the edit-state is applied the same way every time) and not changed by drawing it
  const first = re.rects; await draw(GEO, 0, true); const second = await page.evaluate(() => ({ rects: globalThis.__rects(), v: globalThis.__E.version, rz: JSON.stringify(globalThis.__E.listResizes()), mv: JSON.stringify(globalThis.__E.list()) }));
  ok(sameRects(first, second.rects) && second.v === re.v, 'drawing the slide again gives the same drawing every time (positions, stretch transforms and moves), and drawing never writes to the edit-state (version unchanged)');
  // (b) every committed element is drawn in the frame its record holds: outline = handles = the resized frame, size() reports it
  const chk = await page.evaluate(() => {
    const { sel, sc, hnd } = globalThis.__s; const out = []; const es = globalThis.__E;
    for (const r of es.listResizes()) {
      sel.select(r.id); const s = sel.size(r.id); const n = hnd.nodes.get(r.id); const o = hnd.root.querySelector('.sf-sel'); const geom = globalThis.__corners(n); const hh = globalThis.__hdl();
      const dist = (a, b) => Math.hypot(a.x - b.x, a.y - b.y);
      out.push({ name: sc.byId.get(r.id).name, size: Math.abs(s.w - r.slideFrame.w) < 1e-9 && Math.abs(s.h - r.slideFrame.h) < 1e-9, frame: ['nw', 'ne', 'se', 'sw'].every((k) => dist(geom[k], hh[k]) < 0.05), drawn: Math.abs(dist(geom.nw, geom.ne) - r.slideFrame.w) < 0.05 && Math.abs(dist(geom.nw, geom.sw) - r.slideFrame.h) < 0.05 || hh.nw.w === 0, d: !!o });
    }
    sel.clear(); return out;
  });
  ok(chk.length >= 11 && chk.every((c) => c.size && c.frame && c.drawn && c.d), 'select a restored element: size() is the committed size, the outline with its handles is the resized frame, the drawn element is that frame', chk.filter((c) => !(c.size && c.frame && c.drawn)));
  // (c) a resize that starts from a restored size builds on it: the new record is the TOTAL since the file
  await draw(SEL, 0, true); await page.evaluate(() => globalThis.__newE()); await draw(SEL, 0, true);
  const t1 = await page.evaluate(() => { const a = globalThis.__resizeBy('front-rect', 'se', 40, 20); return { a, rec: globalThis.__E.getResize(globalThis.__s.sel.id), v: globalThis.__E.version }; });
  await draw(SEL, 0, true);
  const t2 = await page.evaluate(() => { const a = globalThis.__resizeBy('front-rect', 'se', 20, 10); return { a, rec: globalThis.__E.getResize(globalThis.__s.sel.id), v: globalThis.__E.version, n: globalThis.__E.listResizes().length }; });
  ok(closeTo(t1.rec.slideFrame.w, 340, 1e-9) && closeTo(t1.rec.slideFrame.h, 220, 1e-9) && closeTo(t2.rec.slideFrame.w, 360, 1e-9) && closeTo(t2.rec.slideFrame.h, 230, 1e-9) && t2.n === 1 && t2.v === t1.v + 1, 'a second resize after a redraw starts from the restored size (300 x 200 -> 340 x 220 -> 360 x 230): one record, the total since the file, version +1', { t1: t1.rec.slideFrame, t2: t2.rec.slideFrame });
  ok(closeTo(t2.rec.origin.w / 12700, 300, 1e-9) && closeTo(t2.rec.extent.w / 12700, 360, 1e-6) && closeTo(t2.rec.extent.h / 12700, 230, 1e-6), 'and the extent is the file\'s 300 x 200 pt scaled to 360 x 230 pt (the origin in the record is still the file\'s, not the previous resize)', t2.rec);
  // (d) back to the file size removes the record
  const t3 = await page.evaluate(() => { const a = globalThis.__resizeBy('front-rect', 'se', -60, -30); return { a, rec: globalThis.__E.getResize(globalThis.__s.sel.id), n: globalThis.__E.listResizes().length, size: globalThis.__E.size, v: globalThis.__E.version }; });
  ok(t3.rec === null && t3.n === 0 && t3.size === 0 && t3.v === t2.v + 1, 'dragging back to exactly the file size removes the record: the element is the file\'s again, nothing is kept for it', t3);
  // (e) another slide and back: records are per slide part
  await page.evaluate(() => globalThis.__newE()); await draw(SEL, 0, true);
  await page.evaluate(() => { globalThis.__resizeBy('front-rect', 'se', 40, 20); });
  const r0 = await rects(); const rec0 = await page.evaluate(() => JSON.stringify(globalThis.__E.listResizes()));
  await draw(SEL, 1, true);
  const s2 = await page.evaluate(() => ({ scaled: globalThis.__s.hnd.root.querySelectorAll('.sf-el[style*="scale("]').length, v: globalThis.__E.version }));
  const resized2 = await page.evaluate(() => { const r = globalThis.__resizeBy(globalThis.__s.sc.elements.filter((e) => e.kind !== 'group' && e.kind !== 'alt')[0].name, 'se', 25, 15); return !!r; });
  await draw(SEL, 2, true); const s3scaled = await page.evaluate(() => globalThis.__s.hnd.root.querySelectorAll('.sf-el[style*="scale("]').length);
  await draw(SEL, 0, true); const back = await rects(); const recBack = await page.evaluate(() => ({ rz: globalThis.__E.listResizes(), v: globalThis.__E.version }));
  ok(s2.scaled === 0 && resized2 && s3scaled === 0, 'slide 2 shows none of slide 1\'s resize (and slide 3 neither): a record belongs to its own slide part', { s2, s3scaled });
  ok(sameRects(r0, back) && recBack.rz.length === 2 && JSON.stringify(recBack.rz.filter((r) => r.slideIndex === 0)) === rec0, 'slide 1 again: front-rect is exactly the size it was left at, the slide-2 resize is kept too (two records), the slide-1 record is byte-identical', { n: recBack.rz.length });
}

// ---- 5. cancellation: nothing reaches the edit-state, an earlier committed resize is exactly what it was --------------------------------------------
{
  const es = () => page.evaluate(() => { const E = globalThis.__E; const id = globalThis.__s.sel.id; return { size: E.size, v: E.version, rz: JSON.stringify(E.listResizes()), mv: JSON.stringify(E.list()), rec: id ? E.getResize(id) : null, dragging: globalThis.__s.sel.dragging, resizing: globalThis.__s.sel.resizing, ph: globalThis.__s.resizes.slice(), sz: id ? globalThis.__s.sel.size(id) : null }; });
  const press = (h, dx, dy) => page.evaluate(({ h, dx, dy }) => { globalThis.__s.resizes.length = 0; return globalThis.__resizeBy('front-rect', h, dx, dy, false); }, { h, dx, dy });
  const fresh = async () => { await page.evaluate(() => globalThis.__newE()); await draw(SEL, 0, true); };
  // a. there is an earlier committed resize (340 x 220): each way of cancelling a second one leaves the record byte-identical, the version unchanged, the frame back at 340 x 220
  await fresh(); await page.evaluate(() => globalThis.__resizeBy('front-rect', 'se', 40, 20)); const kept = await es();
  ok(kept.rec && closeTo(kept.rec.slideFrame.w, 340, 1e-9) && kept.size === 1 && kept.v === 1, 'setup: a completed resize is committed (340 x 220 pt, version 1)', kept.rec);
  const kinds = [
    ['Escape', () => page.keyboard.press('Escape')],
    ['pointercancel', () => fireP('pointercancel', 300, 300)],
    ['a second finger', () => fireP('pointerdown', 100, 100, { pointerId: 2, isPrimary: false, pointerType: 'touch' })],
    ['a right button press', () => fireP('pointerdown', 100, 100, { button: 2, buttons: 2 })],
    ['blur', () => page.evaluate(() => window.dispatchEvent(new Event('blur')))],
    ['select(another element)', () => page.evaluate(() => globalThis.__s.sel.select(globalThis.__s.sc.elements.find((e) => e.name === 'back-rect').id))],
    ['clear()', () => page.evaluate(() => globalThis.__s.sel.clear())],
    ['dispose()', () => page.evaluate(() => globalThis.__s.sel.dispose())],
  ];
  for (const [label, act] of kinds) {
    await fresh(); await page.evaluate(() => globalThis.__resizeBy('front-rect', 'se', 40, 20)); const k0 = await es();
    await press('e', 30, 0); const open = await es(); const wOpen = !!open.sz && Math.abs(open.sz.w - 370) < 1e-9 && open.resizing === 'e';
    await act(); const k1 = await es(); const untouched = k1.rz === k0.rz && k1.v === k0.v && k1.size === k0.size && k1.mv === k0.mv;
    if (label === 'dispose()') { // the controller is gone: the preview is off the slide, the edit-state is what it was and still holds the committed 340 x 220
      const rec = JSON.parse(k1.rz)[0];
      ok(wOpen && untouched && rec && closeTo(rec.slideFrame.w, 340, 1e-9) && closeTo(rec.slideFrame.h, 220, 1e-9), `${label} during a resize: the edit-state is byte-identical to before (record, version ${k0.v}), no write; it still holds the committed 340 x 220 pt`, { wOpen, untouched });
      continue;
    }
    await page.evaluate(() => globalThis.__s.sel.select(globalThis.__s.sc.elements.find((e) => e.name === 'front-rect').id));
    const sz = await page.evaluate(() => globalThis.__s.sel.size(globalThis.__s.sel.id));
    ok(wOpen && untouched && Math.abs(sz.w - 340) < 1e-9 && Math.abs(sz.h - 220) < 1e-9 && k1.dragging === null && JSON.stringify(k1.ph) === '["start","cancel"]', `${label} during a resize: the edit-state is byte-identical to before (record, version ${k0.v}), no write; the frame is back at the committed 340 x 220 pt and start, cancel were reported`, { wOpen, untouched, sz, ph: k1.ph });
  }
  // b. no earlier resize: a cancelled drag creates nothing
  for (const [label, act] of kinds.slice(0, 5)) {
    await fresh(); await press('se', 50, 30); const open = await es(); await act(); const k1 = await es();
    ok(open.sz && open.sz.dw === 50 && k1.size === 0 && k1.v === 0 && k1.rz === '[]' && k1.mv === '[]', `${label} during the first resize of an element: nothing was committed (size 0, version 0)`, { open: open.sz, k1: [k1.size, k1.v] });
  }
  // c. a stray release of another pointer does not commit; the real release does; a flick (release far from the press, no move) commits the total
  await fresh(); await press('se', 20, 10); await fireP('pointerup', 0, 0, { pointerId: 9, isPrimary: false }); let s = await es();
  ok(s.resizing === 'se' && s.v === 0 && s.size === 0, 'the release of some other pointer neither ends the resize nor commits it', s);
  await fireP('pointercancel', 300, 300); s = await es(); ok(s.v === 0 && s.size === 0 && s.sz.dw === 0, 'and the cancel that follows leaves the edit-state empty and the element at the file size', s.sz);
  await fresh(); await page.evaluate(() => globalThis.__s.sel.select(globalThis.__s.sc.elements.find((e) => e.name === 'front-rect').id));
  const c3 = await page.evaluate(() => globalThis.__hdl().se); await fireP('pointerdown', c3.x, c3.y); await fireP('pointerup', c3.x + 60, c3.y + 30); s = await es();
  ok(s.v === 1 && s.rec && closeTo(s.rec.slideFrame.w, 360, 1e-9) && closeTo(s.rec.slideFrame.h, 230, 1e-9) && JSON.stringify(s.ph) === '["start","end"]', 'a release far from the press without a move event commits the total (360 x 230 pt), version 1', s.rec && s.rec.slideFrame);
  // d. a click on a handle (no travel) commits nothing
  await fresh(); await page.evaluate(() => globalThis.__s.sel.select(globalThis.__s.sc.elements.find((e) => e.name === 'front-rect').id)); const h = await page.evaluate(() => globalThis.__hdl().e);
  await fireP('pointerdown', h.x, h.y); await fireP('pointerup', h.x + 1, h.y); s = await es();
  ok(s.v === 0 && s.size === 0, 'a click on a handle without travel resizes nothing and commits nothing', s);
  // e. an edit-state that refuses the resize: the release counts as a cancel (start, cancel), the element goes back, the slide never shows a size the model does not hold
  const ref = await page.evaluate(async (b) => {
    const { parsed, scene } = await globalThis.__load(b); const rn = await import('/js/render/render.js'); const sl = await import('/js/render/selection.js');
    const sc = await scene.buildScene(parsed, 0); globalThis.__s?.sel.dispose();
    const hnd = await rn.renderSlide(sc, document.getElementById('sf-host'), { width: 960 }); const phases = []; let asked = 0; let got = null;
    const sel = sl.attachSelection(hnd, sc, { onResize: (m) => phases.push(m.phase), edits: { forSlide: () => [], resizesForSlide: () => [], textsForSlide: () => [], effectiveText: (e) => e.text, commitText: () => ({ ok: true, edit: null, changed: false }), discardText() {}, commit: () => ({ ok: true }), commitResize: (_s, _i, f) => { asked++; got = f; return { ok: false, reason: 'singular' }; }, discard() {}, discardResize() {} } });
    globalThis.__s = { hnd, sc, sel, parsed, changes: [], moves: [], resizes: [] };
    const id = sc.elements.find((e) => e.name === 'front-rect').id; const r = globalThis.__resizeBy('front-rect', 'se', 40, 20);
    return { phases, asked, got, size: sel.size(id), scaled: hnd.root.querySelectorAll('.sf-el[style*="scale("]').length, dragging: sel.dragging, r };
  }, b64(SEL));
  ok(ref.asked === 1 && JSON.stringify(ref.phases) === '["start","cancel"]' && ref.size.dw === 0 && ref.size.dh === 0 && ref.scaled === 0 && ref.dragging === null && closeTo(ref.got.w, 340, 1e-9) && closeTo(ref.got.tx, 20, 1e-9), 'if the edit-state refuses the resize the release counts as a cancel: the element goes back and the slide never shows a size the model does not hold (the frame it was asked to keep was 340 x 220, centre shifted 20 / 10)', ref);
  // f. an edit-state that accepts: the controller hands over the frame in slide points, and nothing else
  const acc = await page.evaluate(async (b) => {
    const { parsed, scene } = await globalThis.__load(b); const rn = await import('/js/render/render.js'); const sl = await import('/js/render/selection.js');
    const sc = await scene.buildScene(parsed, 0); globalThis.__s?.sel.dispose(); const calls = [];
    const hnd = await rn.renderSlide(sc, document.getElementById('sf-host'), { width: 960 });
    const sel = sl.attachSelection(hnd, sc, { edits: { forSlide: () => [], resizesForSlide: () => [], textsForSlide: () => [], effectiveText: (e) => e.text, commitText: () => ({ ok: true, edit: null, changed: false }), discardText() {}, commit: () => ({ ok: true }), commitResize: (s, i, f) => { calls.push([s === sc, i, f && { ...f }]); return { ok: true }; }, discard() {}, discardResize() {} } });
    globalThis.__s = { hnd, sc, sel, parsed, changes: [], moves: [], resizes: [] }; globalThis.__resizeBy('front-rect', 'se', 40, 20); globalThis.__resizeBy('front-rect', 'se', -40, -20);
    return { calls, id: sc.elements.find((e) => e.name === 'front-rect').id };
  }, b64(SEL));
  ok(acc.calls.length === 2 && acc.calls[0][0] === true && acc.calls[0][1] === acc.id && closeTo(acc.calls[0][2].w, 340, 1e-9) && closeTo(acc.calls[0][2].h, 220, 1e-9) && closeTo(acc.calls[0][2].tx, 20, 1e-9) && closeTo(acc.calls[0][2].ty, 10, 1e-9) && acc.calls[1][2] === null, 'on release the controller asks the edit-state once per completed resize with the scene, the id and the frame { w, h, tx, ty } in slide points (and null when the element is back at the file size)', acc.calls);
}

// ---- 6. size limits, in what is committed (real mouse): 2 pt minimum, four slide sizes maximum, flat axes, group children ------------------------------
{
  const fresh = async (file) => { await page.evaluate(() => globalThis.__newE()); await draw(file, 0, true); };
  const rec = () => page.evaluate(() => globalThis.__E.getResize(globalThis.__s.sel.id));
  const E12 = 12700;
  await fresh(SEL); await page.evaluate(() => globalThis.__resizeBy('front-rect', 'se', -500, -500)); let r = await rec();
  ok(r && r.slideFrame.w === 2 && r.slideFrame.h === 2 && closeTo(r.extent.w, 2 * E12, 1e-6) && closeTo(r.extent.h, 2 * E12, 1e-6) && closeTo(r.position.x, r.origin.x, 1e-6) && closeTo(r.position.y, r.origin.y, 1e-6), 'a corner dragged far inside out commits exactly 2 x 2 pt (extent 25400 EMU), anchored at the top-left: the <a:off> is the file\'s', r && { f: r.slideFrame, e: r.extent, p: r.position, o: r.origin });
  await fresh(SEL); await page.evaluate(() => globalThis.__resizeBy('front-rect', 'e', -400, 0)); r = await rec();
  ok(r && r.slideFrame.w === 2 && r.slideFrame.h === 200 && closeTo(r.extent.h, r.origin.h, 1e-6) && r.scale.sy === 1, 'an edge dragged past the opposite one commits 2 pt there and leaves the other axis exactly as the file has it (scale 1, same extent)', r && { f: r.slideFrame, s: r.scale });
  await fresh(SEL); await page.evaluate(() => globalThis.__resizeBy('front-rect', 'e', 20000, 0)); r = await rec();
  ok(r && closeTo(r.slideFrame.w, 4 * 960, 1e-6) && closeTo(r.extent.w, r.origin.w * (4 * 960 / (r.origin.w / E12)), 1e-3), 'a huge drag commits four slide widths (3840 pt), and the extent is the file\'s extent scaled to it', r && r.slideFrame);
  await fresh(SEL); await page.evaluate(() => globalThis.__resizeBy('h-line', 'se', 40, 30)); r = await rec();
  ok(r && closeTo(r.slideFrame.w, 340, 1e-6) && r.slideFrame.h === 0 && r.scale.sy === 1 && r.extent.h === 0 && r.origin.h === 0 && closeTo(r.extent.w, r.origin.w + 40 * E12, 1e-6) && closeTo(r.position.y, r.origin.y, 1e-9) && Number.isFinite(r.position.x), 'a zero-height line: the commit widens it by 40 pt, its extent height stays 0, its vertical scale is exactly 1 (no NaN, no growth), its <a:off> y is the file\'s', r);
  // group children: the 2 pt minimum is a WORLD size, the extent in the group's own space follows the group's scale
  await fresh(GEO); const kids = {};
  for (const n of ['g07-child', 'g08-child-rot10', 'g09-child']) { const x = await page.evaluate((n) => globalThis.__resizeBy(n, 'se', -500, -500), n); kids[n] = x && (await rec()); }
  ok(Object.values(kids).every((k) => k && k.slideFrame.w === 2 && k.slideFrame.h === 2 && k.extent.w > 0 && k.extent.h > 0 && k.parentId), 'children of rotated, flipped and nested groups dragged far inside out commit 2 x 2 pt in the world, with a positive extent in their group\'s space and a parent id', kids);
  const vv = await page.evaluate((b) => globalThis.__verifyFrames(b, 0), b64(GEO));
  ok(vv.edited === 3 && vv.rows.filter((x) => x.resized).every((x) => x.err < 16 && x.rotErr < 1e-6 && x.flipSame) && vv.rows.filter((x) => x.resized).length === 3, 'and a Scene Model rebuilt with the committed <a:off> / <a:ext> of those three draws them at exactly 2 x 2 pt, same centre rule, same rotation and mirroring', vv.rows.filter((x) => x.resized));
  await page.evaluate(() => globalThis.__s.sel.dispose());
}

// ---- 7. the original XML, the Scene Model and the file stay untouched (every fixture and real deck), and the resize comes back from the edit-state ------
{
  const decks = [GEO, SEL, path.join(FX, 'p31-shapes.pptx'), path.join(FX, 'p31-text.pptx'), path.join(FX, 'p31-arabic.pptx'), path.join(FX, 'p31-media.pptx'), path.join(FX, 'p31-unsupported.pptx'), path.join(FX, 'p31-template-decor.pptx'), path.join(FX, 'p22-geometry.pptx'), ...realFiles];
  for (const f of decks) {
    const r = await page.evaluate(async ({ b }) => {
      const { parsed, scene } = await globalThis.__load(b); const rn = await import('/js/render/render.js'); const sl = await import('/js/render/selection.js'); const em = await import('/js/model/edit-state.js');
      const ser = (d) => new XMLSerializer().serializeToString(d); const edits = em.createEditState();
      const zip0 = await globalThis.__zipKey(parsed.zip); const orig = await globalThis.__zipKeyOf(b);
      const out = { slides: 0, committed: 0, restored: 0, mutated: 0, worldChanged: 0, dirty: 0, bad: 0, errors: [], zipSame: zip0 === orig, zipAfter: true };
      for (let i = 0; i < parsed.slides.length; i++) {
        try {
          const sc = await scene.buildScene(parsed, i); const docs = [...new Set(sc.all.map((e) => e.el.ownerDocument))]; const before = docs.map(ser);
          const w0 = JSON.stringify(sc.all.map((e) => [e.world, e.xfrm]));
          const hnd = await rn.renderSlide(sc, document.getElementById('sf-host'), { width: 960 }); const html0 = hnd.root.outerHTML;
          let sel = sl.attachSelection(hnd, sc, { edits }); const R = hnd.root.getBoundingClientRect();
          const ev = (t, x, y) => hnd.root.dispatchEvent(new PointerEvent(t, { bubbles: true, cancelable: true, pointerType: 'mouse', isPrimary: true, pointerId: 1, button: 0, buttons: t === 'pointerup' ? 0 : 1, clientX: R.left + x, clientY: R.top + y }));
          let hit = null;
          for (const e of sc.elements) {
            if (e.kind === 'group' || e.kind === 'alt' || !hnd.nodes.has(e.id) || !sel.select(e.id)) continue;
            const hd = hnd.root.querySelector('.sf-sel .sf-hdl[data-h="se"]'); const r0 = hd && hd.getBoundingClientRect(); if (!r0 || !(r0.width > 0)) continue;
            const x = r0.left + r0.width / 2 - R.left; const y = r0.top + r0.height / 2 - R.top;
            ev('pointerdown', x, y); ev('pointermove', x + 21, y + 13); ev('pointerup', x + 21, y + 13);
            if (edits.hasResize(e.id)) { hit = e.id; break; }
          }
          if (hit) {
            out.committed++; const rec = edits.getResize(hit); const fr = edits.frame(hit);
            if (!(rec && [rec.extent.w, rec.extent.h, rec.position.x, rec.position.y, fr.x, fr.y, fr.w, fr.h].every(Number.isFinite) && rec.extent.w >= 0 && rec.extent.h >= 0)) out.bad++;
            sel.dispose(); if (hnd.root.outerHTML !== html0) out.dirty++;
            sel = sl.attachSelection(hnd, sc, { edits }); // a redraw of the same slide: the resize is drawn again
            const s = sel.size(hit);
            if (Math.abs(s.w - rec.slideFrame.w) < 1e-9 && Math.abs(s.h - rec.slideFrame.h) < 1e-9 && (s.dw || s.dh) && /scale\(/.test(hnd.nodes.get(hit).style.transform)) out.restored++;
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
    ok(r.errors.length === 0 && r.slides > 0 && r.committed > 0 && r.restored === r.committed && r.bad === 0, `${nm}: ${r.committed} of ${r.slides} slide(s) committed a resize and drew it again from the edit-state without errors (finite extent, position and frame)`, r);
    ok(r.dirty === 0 && r.mutated === 0 && r.worldChanged === 0 && r.zipSame && r.zipAfter, `${nm}: slide / layout / master XML, every part of the file (SHA-256 per part, before and after), every world frame and xfrm are identical, and the drawn DOM is clean after dispose`, r);
  }
  // code guards
  const code = readFileSync(path.join(ROOT, 'js/render/selection.js'), 'utf8').replace(/\/\/.*$/gm, '');
  const fin = code.slice(code.indexOf('function finishDrag'), code.indexOf('function select('));
  ok(/edits\.commitResize\(scene, dr\.id, sizes\.get\(dr\.id\) \|\| null\)/.test(fin) && (code.match(/commitResize/g) || []).length === 1, 'selection.js commits a resize in exactly one place: finishDrag(), with the frame the preview drew (the drawing code never reaches the edit-state)');
  ok(!/\.xfrm\b|\.localShift\b|\.extent\b|\.position\.[xy]\b|\.origin\.[wh]\b/.test(code), 'selection.js hands the frame to the edit-state and never computes or writes an extent or a local coordinate itself');
  ok(!/^const (MIN_PT|FLAT_PT|MAX_TIMES)\b/m.test(code) && /import \{ MIN_PT, FLAT_PT, MAX_TIMES, SAME_PT \} from '\.\.\/model\/resize-limits\.js'/.test(code), 'the size limits exist once (js/model/resize-limits.js) and the drag and the edit-state both use them');
  const es = readFileSync(path.join(ROOT, 'js/model/edit-state.js'), 'utf8').replace(/\/\/.*$/gm, '');
  ok(/MIN_PT/.test(es) && /MAX_TIMES/.test(es) && /FLAT_PT/.test(es), 'the edit-state checks the same limits itself (a frame the drag could never produce is not kept)');
  const rd = readFileSync(path.join(ROOT, 'js/ui/rendered.js'), 'utf8').replace(/\/\/.*$/gm, '');
  ok(!/localStorage|indexedDB|store\.templates|\.update\(/.test(rd), 'the Rendered tab still keeps the edit-state in memory only: no storage call was added');
}

// ---- 8. the real Rendered tab (mouse): a resize survives redraws, slides, tabs and zoom; Escape cancels; nothing is stored; nothing outlives the opened deck ----
{
  const ui = await ctx.newPage(); watch(ui, 'resize-commit-ui: ');
  await ui.setViewportSize({ width: 1280, height: 900 });
  await openRendered(ui, 'Resize commit deck');
  const store0 = await ui.evaluate(async () => ({ ls: JSON.stringify({ ...localStorage }), dbs: ((await indexedDB.databases?.()) || []).map((d) => d.name).sort().join() }));
  const info = () => ui.evaluate(() => {
    const root = document.querySelector('.rendered-host .sf-slide'); const R = root.getBoundingClientRect(); const els = {};
    for (const n of root.querySelectorAll('.sf-el')) { const r = n.getBoundingClientRect(); els[n.dataset.name] = { l: r.left - R.left, t: r.top - R.top, w: r.width, h: r.height, scaled: /scale\(/.test(n.style.transform), tr: n.style.translate }; }
    const o = root.querySelector('.sf-sel'); const hd = o && o.querySelector('.sf-hdl[data-h="se"]'); const q = hd && hd.getBoundingClientRect();
    return { els, se: q && [q.left + q.width / 2, q.top + q.height / 2], status: document.querySelector('.sf-sel-status')?.textContent || '', slide: root.dataset.slide, scaled: root.querySelectorAll('.sf-el[style*="scale("]').length, outlines: root.querySelectorAll('.sf-sel').length };
  });
  const sc = (await slideBox(ui)).s;
  await ui.mouse.click(...(await at(ui, 450, 300)));
  const a = await info(); const fr0 = a.els['front-rect'];
  await ui.mouse.move(a.se[0], a.se[1]); await ui.mouse.down(); await ui.mouse.move(a.se[0] + 30 * sc, a.se[1] + 20 * sc, { steps: 4 }); await ui.mouse.move(a.se[0] + 60 * sc, a.se[1] + 40 * sc, { steps: 4 }); await ui.mouse.up();
  const b = await info(); const fr1 = b.els['front-rect'];
  ok(Math.abs(fr1.w - fr0.w - 60 * sc) <= 1 && Math.abs(fr1.h - fr0.h - 40 * sc) <= 1 && /Resized in the preview to 360 × 240 pt/.test(b.status), 'real mouse: the element is 60 / 40 pt bigger after the release (360 x 240 pt)', { fr0, fr1, status: b.status });
  // a drag on the body moves it as well: both are kept
  const sb = await slideBox(ui); const bx = sb.x + fr1.l + fr1.w / 2; const by = sb.y + fr1.t + fr1.h / 2;
  await ui.mouse.move(bx, by); await ui.mouse.down(); await ui.mouse.move(bx + 20 * sc, by + 10 * sc, { steps: 4 }); await ui.mouse.up();
  const m = await info(); const fr2 = m.els['front-rect'];
  ok(Math.abs(fr2.l - fr1.l - 20 * sc) <= 1 && Math.abs(fr2.w - fr1.w) <= 0.5, 'a drag on the body moves the resized element and keeps its size', { fr1, fr2 });
  // the edit-state of the open deck holds both
  // another slide and back, three times, with a resize on slide 2 as well
  await ui.click('.rendered .film-item >> nth=1'); await ui.waitForSelector('.rendered-host .sf-slide[data-slide="1"] .sf-el');
  const s2a = await info();
  ok(s2a.scaled === 0 && !s2a.els['front-rect'], 'slide 2 does not show slide 1\'s resize', Object.keys(s2a.els));
  await ui.mouse.click(...(await at(ui, 150, 150))); const s2s = await info();
  await ui.mouse.move(s2s.se[0], s2s.se[1]); await ui.mouse.down(); await ui.mouse.move(s2s.se[0] + 25 * sc, s2s.se[1] + 15 * sc, { steps: 5 }); await ui.mouse.up();
  const s2b = await info();
  ok(s2b.scaled === 1 && /Resized in the preview/.test(s2b.status), 'a resize on slide 2 works the same', s2b.status);
  for (let k = 0; k < 3; k++) {
    await ui.click('.rendered .film-item >> nth=0'); await ui.waitForSelector('.rendered-host .sf-slide[data-slide="0"] .sf-el');
    const c1 = await info(); const f = c1.els['front-rect'];
    ok(Math.abs(f.l - fr2.l) <= 0.01 && Math.abs(f.t - fr2.t) <= 0.01 && Math.abs(f.w - fr2.w) <= 0.01 && Math.abs(f.h - fr2.h) <= 0.01 && c1.scaled === 1 && c1.outlines === 0, `slide 1 again (round ${k + 1}): front-rect is exactly the size and place it was left at, it is the only stretched element`, { fr2, f, scaled: c1.scaled });
    await ui.click('.rendered .film-item >> nth=1'); await ui.waitForSelector('.rendered-host .sf-slide[data-slide="1"] .sf-el');
    const c2 = await info(); const s = c2.els[Object.keys(c2.els).find((n) => c2.els[n].scaled)];
    ok(c2.scaled === 1 && s && Math.abs(s.w - (s2b.els[Object.keys(s2b.els).find((n) => s2b.els[n].scaled)].w)) <= 0.01, `slide 2 again (round ${k + 1}): its resized element is exactly as it was left`);
  }
  // click the resized element on a redrawn slide: the status line says what the edit-state holds
  await ui.click('.rendered .film-item >> nth=0'); await ui.waitForSelector('.rendered-host .sf-slide[data-slide="0"] .sf-el');
  const cur = await info(); const sb2 = await slideBox(ui);
  await ui.mouse.click(sb2.x + (cur.els['front-rect'].l + cur.els['front-rect'].w / 2), sb2.y + (cur.els['front-rect'].t + cur.els['front-rect'].h / 2));
  const sel1 = await info();
  ok(/Moved in the preview by 20 pt, 10 pt \(not saved\)\. Resized in the preview to 360 × 240 pt \(not saved\)/.test(sel1.status), 'select it again on the redrawn slide: the status line shows the kept move and the kept size', sel1.status);
  // the other tabs and back, and the filmstrip of the Slides tab
  await ui.click('role=tab[name="Design"]'); await ui.click('role=tab[name="Slides"]'); await ui.click('.preview:not(.rendered) .film-item >> nth=0'); await ui.click('role=tab[name="Design"]'); await ui.click('role=tab[name="Rendered"]'); await ui.waitForSelector('.rendered-host .sf-slide[data-slide="0"] .sf-el');
  const t1 = await info(); const ft = t1.els['front-rect'];
  ok(Math.abs(ft.w - fr2.w) <= 0.01 && Math.abs(ft.h - fr2.h) <= 0.01 && Math.abs(ft.l - fr2.l) <= 0.01 && t1.slide === '0', 'Slides, Design and Rendered tab round trips (and choosing the slide in the Slides tab) keep the resize and the move', { fr2, ft });
  // Escape in the middle of a second resize: back to the kept size, and that is what comes back after a slide change
  const e0 = await info();
  await ui.mouse.move(e0.se[0], e0.se[1]); await ui.mouse.down(); await ui.mouse.move(e0.se[0] + 40 * sc, e0.se[1] + 30 * sc, { steps: 4 });
  const e1 = await info(); await ui.keyboard.press('Escape'); await ui.mouse.up(); const e2 = await info();
  ok(Math.abs(e1.els['front-rect'].w - fr2.w - 40 * sc) <= 1 && Math.abs(e2.els['front-rect'].w - fr2.w) <= 0.01 && Math.abs(e2.els['front-rect'].h - fr2.h) <= 0.01, 'Escape during a second resize puts the element back to the kept 360 x 240 pt (the preview was 40 pt wider)', { fr2, e1: e1.els['front-rect'], e2: e2.els['front-rect'] });
  await ui.click('.rendered .film-item >> nth=1'); await ui.waitForSelector('.rendered-host .sf-slide[data-slide="1"] .sf-el'); await ui.click('.rendered .film-item >> nth=0'); await ui.waitForSelector('.rendered-host .sf-slide[data-slide="0"] .sf-el');
  const e3 = await info(); ok(Math.abs(e3.els['front-rect'].w - fr2.w) <= 0.01 && Math.abs(e3.els['front-rect'].h - fr2.h) <= 0.01, 'and after another slide and back it is still the kept size: the cancelled drag never reached the edit-state', { e3: e3.els['front-rect'] });
  // zoom: the size is kept in slide points
  await ui.setViewportSize({ width: 820, height: 900 });
  await ui.waitForFunction(() => { const r = document.querySelector('.rendered-host .sf-slide'); const f = document.querySelector('.sf-frame'); return r && Math.abs(r.getBoundingClientRect().width - f.clientWidth) <= 1 && f.clientWidth < 800; });
  const z = await info(); const sc2 = (await slideBox(ui)).s;
  ok(Math.abs(z.els['front-rect'].w / sc2 - fr2.w / sc) <= 0.05 && Math.abs(z.els['front-rect'].h / sc2 - fr2.h / sc) <= 0.05, 'after the window is resized the kept size is still 360 x 240 pt (the resize is kept in slide points)', { z: z.els['front-rect'], sc2, sc });
  await ui.setViewportSize({ width: 1280, height: 900 });
  const store1 = await ui.evaluate(async () => ({ ls: JSON.stringify({ ...localStorage }), dbs: ((await indexedDB.databases?.()) || []).map((d) => d.name).sort().join() }));
  ok(store1.ls === store0.ls && store1.dbs === store0.dbs, 'resizing elements wrote nothing to localStorage and created no IndexedDB database', { store0, store1 });
  // closing the deck ends the edit-state: opening it again shows the file
  await ui.click('.modal [aria-label="Close"]'); await ui.waitForSelector('.modal', { state: 'detached' });
  await ui.click('.tpl:has-text("Resize commit deck") >> text=Preview'); await ui.waitForSelector('.film-item');
  await ui.click('role=tab[name="Rendered"]'); await ui.waitForSelector('.rendered-host .sf-slide .sf-el');
  const r1 = await info();
  ok(r1.scaled === 0 && Object.values(r1.els).every((e) => !e.tr) && Math.abs(r1.els['front-rect'].w - fr0.w) <= 0.01 && Math.abs(r1.els['front-rect'].h - fr0.h) <= 0.01, 'closing the preview and opening the template again shows every element at its file size and position: the edit-state is temporary and ends with the opened deck', r1.els['front-rect']);
  await ui.close();
}

// ---- 9. a finger: a resize by touch is committed and survives a slide change ------------------------------------------------------------------------
{
  const tctx = await browser.newContext({ viewport: { width: 1280, height: 900 }, hasTouch: true, deviceScaleFactor: 2 });
  const ui = await tctx.newPage(); watch(ui, 'resize-commit-touch: ');
  await openRendered(ui, 'Resize commit touch deck');
  await ui.setViewportSize({ width: 390, height: 844 });
  await ui.waitForFunction(() => document.querySelector('.rendered-host .sf-slide').getBoundingClientRect().width < 380);
  await ui.evaluate(() => document.querySelector('.rendered-host .sf-slide').scrollIntoView({ block: 'center' }));
  const cdp = await tctx.newCDPSession(ui); const sc = (await slideBox(ui)).s;
  const info = () => ui.evaluate(() => { const root = document.querySelector('.rendered-host .sf-slide'); const n = [...root.querySelectorAll('.sf-el')].find((x) => x.dataset.name === 'front-rect'); const r = n.getBoundingClientRect(); const hd = root.querySelector('.sf-sel .sf-hdl[data-h="se"]'); const q = hd && hd.getBoundingClientRect(); return { w: r.width, h: r.height, se: q && [q.left + q.width / 2, q.top + q.height / 2], scaled: root.querySelectorAll('.sf-el[style*="scale("]').length, scrollY: window.scrollY }; });
  await ui.touchscreen.tap(...(await at(ui, 450, 300)));
  const a = await info();
  const drag = async (x, y, dx, dy) => { await cdp.send('Input.dispatchTouchEvent', { type: 'touchStart', touchPoints: [{ x, y, id: 1 }] }); for (let i = 1; i <= 4; i++) await cdp.send('Input.dispatchTouchEvent', { type: 'touchMove', touchPoints: [{ x: x + (dx * i) / 4, y: y + (dy * i) / 4, id: 1 }] }); await cdp.send('Input.dispatchTouchEvent', { type: 'touchEnd', touchPoints: [] }); await ui.waitForTimeout(150); };
  await drag(a.se[0], a.se[1], 30 * sc, 20 * sc);
  const b = await info();
  await ui.click('.rendered .film-item >> nth=1'); await ui.waitForSelector('.rendered-host .sf-slide[data-slide="1"] .sf-el');
  await ui.click('.rendered .film-item >> nth=0'); await ui.waitForSelector('.rendered-host .sf-slide[data-slide="0"] .sf-el');
  const c = await info();
  ok(Math.abs(b.w - a.w - 30 * sc) <= 1 && Math.abs(b.h - a.h - 20 * sc) <= 1 && Math.abs(c.w - b.w) <= 0.01 && Math.abs(c.h - b.h) <= 0.01 && c.scaled === 1, 'a finger on the south-east handle resizes the element (+30 / +20 pt), and after another slide and back it is still that size', { a, b, c });
  await tctx.close();
}


// ======================================================================================================
section('integration hardening (Phase 3.2.4-C): combined move + resize, clamp, cancellation, redraw, state consistency');
await page.evaluate(() => {
  /** Every drawn element against the edit-state: the offset and size the controller shows are exactly what the edit-state holds (or the file's), and the DOM says the same (translate iff moved, scale iff resized). Returns the list of mismatches. */
  globalThis.__consistency = () => {
    const { sel, sc, hnd } = globalThis.__s; const E = globalThis.__E; const bad = [];
    for (const e of sc.elements) {
      if (e.kind === 'group' || e.kind === 'alt' || !hnd.nodes.has(e.id)) continue;
      const mv = E.get(e.id); const rz = E.getResize(e.id); const o = sel.offset(e.id); const s = sel.size(e.id); const eo = mv ? mv.slideDelta : { dx: 0, dy: 0 };
      if (Math.abs(o.dx - eo.dx) > 1e-9 || Math.abs(o.dy - eo.dy) > 1e-9) bad.push([e.name, 'offset', o, eo]);
      const ew = rz ? rz.slideFrame.w : e.world.w / 12700; const eh = rz ? rz.slideFrame.h : e.world.h / 12700;
      if (Math.abs(s.w - ew) > 1e-9 || Math.abs(s.h - eh) > 1e-9) bad.push([e.name, 'size', s, [ew, eh]]);
      const n = hnd.nodes.get(e.id); if (!!n.style.translate !== !!mv) bad.push([e.name, 'translate-in-dom', !!n.style.translate, !!mv]); if (/scale\(/.test(n.style.transform) !== !!rz) bad.push([e.name, 'scale-in-dom', /scale\(/.test(n.style.transform), !!rz]);
    }
    return bad;
  };
  globalThis.__editSnap = () => { const g = globalThis; const E = g.__E; const id = g.__s.sel.id; return { mv: JSON.stringify(E.list()), rz: JSON.stringify(E.listResizes()), v: E.version, size: E.size, off: id && g.__s.sel.offset(id), sz: id && g.__s.sel.size(id), dragging: g.__s.sel.dragging }; };
  /** A deck whose front-rect (file frame 200..500 x 140..340 pt) is moved to (x, y) EMU in the file copy. */
  globalThis.__variant = async (b64, x, y) => {
    const zip = await JSZip.loadAsync(Uint8Array.from(atob(b64), (c) => c.charCodeAt(0))); const f = zip.file('ppt/slides/slide1.xml'); let xml = await f.async('string');
    const a = '<a:off x="2540000" y="1778000"/>'; if (!xml.includes(a)) throw new Error('front-rect off not found'); zip.file('ppt/slides/slide1.xml', xml.replace(a, `<a:off x="${x}" y="${y}"/>`));
    return zip.generateAsync({ type: 'base64' });
  };
});
const newE = () => page.evaluate(() => globalThis.__newE());
const snapE = () => page.evaluate(() => globalThis.__editSnap());
const consistent = () => page.evaluate(() => globalThis.__consistency());

// ---- 1. the move clamp: an element whose centre is already outside the slide is not snapped back (it used to jump by up to the whole distance) -----------
{
  // (a) a resize stretched the centre off the slide: east handle +1500 pt moves the centre from 350 to 1100 pt (the slide is 960 wide)
  await newE(); await draw(SEL, 0, true);
  const r = await page.evaluate(() => {
    const g = globalThis; const out = {}; const rs = g.__resizeBy('front-rect', 'e', 1500, 0); out.size = rs && rs.w; const id = g.__s.sel.id;
    g.__moveBy('front-rect', 10, 0); out.small = { off: g.__s.sel.offset(id), rec: g.__E.get(id) };
    g.__moveBy('front-rect', -300, 0); out.left = g.__s.sel.offset(id);
    g.__moveBy('front-rect', -1200, 0); out.farLeft = g.__s.sel.offset(id);
    g.__moveBy('front-rect', 50, 0); out.back = g.__s.sel.offset(id);
    return out;
  });
  ok(r.size === 1800 && r.small.off.dx === 0 && r.small.rec === null, 'a resize left the centre outside the slide (1100 of 960 pt): a 10 pt drag further out leaves it where it was, nothing is committed (it used to be thrown to -140 pt, the other way, and committed)', r.small);
  ok(Math.abs(r.left.dx - -300) < 1e-6, 'dragging it back in by 300 pt moves it exactly 300 pt (centre 800, inside the slide)', r.left);
  ok(Math.abs(r.farLeft.dx - -1100) < 1e-6, 'dragging it far past the left edge still stops with the centre on the slide (centre at 0 pt: the clamp still holds on the way in)', r.farLeft);
  ok(Math.abs(r.back.dx - -1050) < 1e-6, 'and from there it follows the drag again (+50 pt)', r.back);
  ok((await consistent()).length === 0, 'drawn offsets and sizes equal the edit-state after all of that');
  // (b) the same without any resize: the FILE has the centre outside (a copy of the deck with front-rect moved to x = 905 pt, y = 449 pt: centre 1055 / 549 of 960 x 540, a corner still on the slide)
  const vb = await page.evaluate(({ b }) => globalThis.__variant(b, 11500000, 5700000), { b: b64(SEL) });
  await newE(); await page.evaluate(({ b }) => globalThis.__draw(b, 0, true), { b: vb });
  const v = await page.evaluate(() => {
    const g = globalThis; const id = g.__s.sc.elements.find((e) => e.name === 'front-rect').id; const o = {}; const W = g.__s.sc.byId.get(id).world; o.centre = [W.cx / 12700, W.cy / 12700];
    g.__moveBy('front-rect', 10, 10); o.small = g.__s.sel.offset(id);
    g.__moveBy('front-rect', 500, 500); o.further = g.__s.sel.offset(id);
    g.__moveBy('front-rect', -100, -200); o.inward = g.__s.sel.offset(id);
    return o;
  });
  ok(v.centre[0] > 960 && v.centre[1] > 540 && v.small.dx === 0 && v.small.dy === 0, 'an element the file places with its centre outside the slide in both axes: a 10 / 10 pt drag further out leaves it where it was (no snap back by 95 / 9 pt in either axis)', v);
  ok(Math.abs(v.further.dx) < 1e-6 && Math.abs(v.further.dy) < 1e-6, 'a long drag further out changes nothing either: it cannot go further out than where the drag started', v.further);
  ok(Math.abs(v.inward.dx - -100) < 1e-6 && Math.abs(v.inward.dy - -200) < 1e-6, 'and it can always be dragged inward: -100 / -200 pt moves it exactly that far', v.inward);
  // (c) elements that start inside keep the 3.2.3 behaviour: the centre stops on the slide edge
  await newE(); await draw(SEL, 0, true);
  const c = await page.evaluate(() => { globalThis.__moveBy('front-rect', 2000, 2000); const id = globalThis.__s.sel.id; return globalThis.__s.sel.offset(id); });
  ok(Math.abs(c.dx - (960 - 350)) < 1e-6 && Math.abs(c.dy - (540 - 240)) < 1e-6, 'an element that starts inside still stops with its centre on the slide edge (offset 610 / 300 pt)', c);
}

// ---- 2. the preview and the edit-state agree on "back at the file frame" (a release a hair away from it left a stretched element the edit-state did not hold) ----
{
  const drive = (off) => page.evaluate((off) => { const g = globalThis; g.__s.sel.select(g.__s.sc.elements.find((e) => e.name === 'front-rect').id); const R = g.__s.hnd.root.getBoundingClientRect(); const hd = g.__s.hnd.root.querySelector('.sf-sel .sf-hdl[data-h="se"]').getBoundingClientRect(); const x = hd.left + hd.width / 2 - R.left; const y = hd.top + hd.height / 2 - R.top; g.__ptr('pointerdown', x, y); g.__ptr('pointermove', x + 30, y + 30); g.__ptr('pointermove', x + off, y); g.__ptr('pointerup', x + off, y); const id = g.__s.sel.id; return { drawn: g.__s.sel.size(id), rec: g.__E.getResize(id), scaled: g.__s.hnd.root.querySelectorAll('.sf-el[style*="scale("]').length, v: g.__E.version, bad: g.__consistency() }; }, off);
  await newE(); await draw(SEL, 0, true); let r = await drive(5e-7);
  ok(r.drawn.dw === 0 && r.drawn.dh === 0 && r.scaled === 0 && r.rec === null && r.v === 0 && r.bad.length === 0, 'released 5e-7 pt from the file frame: the slide shows the file frame and the edit-state holds nothing (they both treat it as the file frame)', r);
  await newE(); await draw(SEL, 0, true); r = await drive(5e-3);
  ok(Math.abs(r.drawn.dw - 5e-3) < 1e-9 && r.scaled === 1 && r.rec && r.rec.slideFrame.w === r.drawn.w && r.v === 1 && r.bad.length === 0, 'released 0.005 pt away: a real (tiny) resize, drawn and committed alike', r);
  const limits = await page.evaluate(async () => { const m = await import('/js/model/resize-limits.js'); return m.SAME_PT; });
  ok(limits === 1e-6, 'the tolerance is one shared constant (SAME_PT, 1e-6 pt)');
}

// ---- 3. a move FIRST, then a resize, over p31-geometry (the order 3.2.4-B did not check): the composite frames rebuild exactly --------------------------
{
  await newE(); await draw(GEO, 0, true);
  const names = await page.evaluate(() => globalThis.__s.sc.elements.filter((e) => e.kind !== 'group' && e.kind !== 'alt').map((e) => e.name));
  const HS = ['nw', 'se', 'n', 'w', 's', 'ne', 'e', 'sw']; let both = 0;
  for (const [i, n] of names.entries()) { const moved = await page.evaluate((n) => globalThis.__moveBy(n, 23, 11), n); const rz = await page.evaluate(({ n, h }) => globalThis.__resizeBy(n, h, 19, 13), { n, h: HS[i % HS.length] }); if (moved && rz && (rz.dw || rz.dh)) both++; }
  const v = await page.evaluate((b) => globalThis.__verifyFrames(b, 0), b64(GEO));
  const touched = v.rows.filter((r) => r.moved || r.resized);
  ok(both >= 9 && v.rows.filter((r) => r.moved && r.resized).length >= 9, `${v.rows.filter((r) => r.moved && r.resized).length} elements (plain, rotated, flipped, group children, nested groups, connectors, text) moved first and resized afterwards`, { both });
  ok(touched.every((r) => r.err < 16 && r.rotErr < 1e-6 && r.flipSame), 'a Scene Model rebuilt with the composite <a:off> / <a:ext> of every one of them equals the drawn frame (move, then resize from the moved position), within rounding to whole EMU', touched.filter((r) => !(r.err < 16)).map((r) => [r.name, r.err]));
  ok((await consistent()).length === 0, 'and the slide shows exactly what the edit-state holds for every element');
  await page.evaluate(() => globalThis.__s.sel.dispose());
}

// ---- 4. cancellation with BOTH a committed move and a committed resize: each way of cancelling a move or a resize leaves both records, the version and what is drawn untouched ----
{
  const setup = async () => { await newE(); await draw(SEL, 0, true); await page.evaluate(() => { globalThis.__resizeBy('front-rect', 'se', 40, 20); globalThis.__moveBy('front-rect', 30, 10); }); return snapE(); };
  const acts = [
    ['Escape', () => page.keyboard.press('Escape')], ['pointercancel', () => fireP('pointercancel', 300, 300)],
    ['a second finger', () => fireP('pointerdown', 100, 100, { pointerId: 2, isPrimary: false, pointerType: 'touch' })], ['a right button press', () => fireP('pointerdown', 100, 100, { button: 2, buttons: 2 })],
    ['blur', () => page.evaluate(() => window.dispatchEvent(new Event('blur')))], ['select(another element)', () => page.evaluate(() => globalThis.__s.sel.select(globalThis.__s.sc.elements.find((e) => e.name === 'back-rect').id))], ['clear()', () => page.evaluate(() => globalThis.__s.sel.clear())],
  ];
  for (const kind of ['move', 'resize']) for (const [label, act] of acts) {
    const a0 = await setup();
    await page.evaluate((kind) => { globalThis.__s.moves.length = 0; globalThis.__s.resizes.length = 0; if (kind === 'move') globalThis.__moveBy('front-rect', 50, 50, false); else globalThis.__resizeBy('front-rect', 'nw', -20, -20, false); }, kind);
    const open = await snapE(); await act(); await page.evaluate(() => globalThis.__s.sel.select(globalThis.__s.sc.elements.find((e) => e.name === 'front-rect').id)); const a1 = await snapE();
    const ph = await page.evaluate((kind) => (kind === 'move' ? globalThis.__s.moves : globalThis.__s.resizes).slice(), kind);
    ok(open.dragging && a1.mv === a0.mv && a1.rz === a0.rz && a1.v === a0.v && JSON.stringify(a1.off) === JSON.stringify(a0.off) && JSON.stringify(a1.sz) === JSON.stringify(a0.sz) && JSON.stringify(ph) === '["start","cancel"]' && (await consistent()).length === 0, `${label} during a ${kind} of an element that was moved and resized: both committed records, the version (${a0.v}), the drawn offset and size are exactly as before; start, cancel reported`, { a0: [a0.v, a0.off, a0.sz], a1: [a1.v, a1.off, a1.sz], ph });
  }
}

// ---- 5. undo and zero: moving back to exactly the file position keeps the resize; reset() undoes both and leaves the DOM byte-identical ------------------
{
  await newE(); await draw(SEL, 0, true);
  await page.evaluate(() => { globalThis.__resizeBy('front-rect', 'se', 40, 20); globalThis.__moveBy('front-rect', 30, 10); globalThis.__moveBy('front-rect', -30, -10); });
  let s = await snapE(); await draw(SEL, 0, true); const re = await page.evaluate(() => ({ scaled: globalThis.__s.hnd.root.querySelectorAll('.sf-el[style*="scale("]').length, translated: [...globalThis.__s.hnd.root.querySelectorAll('.sf-el')].filter((n) => n.style.translate).length, bad: globalThis.__consistency() }));
  ok(JSON.parse(s.mv).length === 0 && JSON.parse(s.rz).length === 1 && s.size === 1 && re.scaled === 1 && re.translated === 0 && re.bad.length === 0, 'moving a resized element back to exactly its file position removes the move record and keeps the resize; a redraw shows the resize only', { s: [s.mv, s.size], re });
  await newE(); await draw(SEL, 0, true);
  const pre = await page.evaluate(() => { globalThis.__s.sel.select(globalThis.__s.sc.elements.find((e) => e.name === 'front-rect').id); return globalThis.__s.hnd.root.outerHTML; });
  await page.evaluate(() => { globalThis.__resizeBy('front-rect', 'se', 40, 20); globalThis.__moveBy('front-rect', 30, 10); globalThis.__s.sel.reset(); });
  const after = await page.evaluate(() => ({ html: globalThis.__s.hnd.root.outerHTML, size: globalThis.__E.size, v: globalThis.__E.version }));
  ok(after.html === pre && after.size === 0 && after.v >= 2, 'reset() on a selected, moved and resized element: the DOM (node and outline) is byte-identical to the file drawing and the edit-state is empty', { same: after.html === pre, size: after.size });
}

// ---- 6. seeded random sequences of moves, resizes, cancelled drags, redraws, selections and resets: what is drawn is always what the edit-state holds ------
{
  const res = await page.evaluate(async ({ b }) => {
    const g = globalThis; const out = []; await g.__newE(); await g.__draw(b, 0, true); const names = g.__s.sc.elements.filter((e) => e.kind !== 'group' && e.kind !== 'alt').map((e) => e.name); // names of p31-geometry (the slide drawn before was another deck)
    const HS = ['nw', 'n', 'ne', 'e', 'se', 's', 'sw', 'w'];
    const rect = () => JSON.stringify(Object.fromEntries(g.__s.sc.elements.filter((e) => g.__s.hnd.nodes.has(e.id)).map((e) => { const r = g.__s.hnd.nodes.get(e.id).getBoundingClientRect(); return [e.name, [Math.round(r.left * 100), Math.round(r.top * 100), Math.round(r.width * 100), Math.round(r.height * 100)]]; })));
    const keydown = (key) => document.dispatchEvent(new KeyboardEvent('keydown', { key, bubbles: true, cancelable: true }));
    for (const seed of [11, 202, 3003, 40004]) {
      await g.__newE(); await g.__draw(b, 0, true); let x = seed; const rnd = () => { x = (x * 1664525 + 1013904223) % 4294967296; return x / 4294967296; }; const pick = (a) => a[Math.floor(rnd() * a.length)]; const hist = [];
      for (let i = 0; i < 45; i++) {
        const n = pick(names); const k = rnd(); let op;
        if (k < 0.28) { op = 'move'; g.__moveBy(n, Math.round(rnd() * 160 - 80), Math.round(rnd() * 160 - 80)); }
        else if (k < 0.56) { op = 'resize'; g.__resizeBy(n, pick(HS), Math.round(rnd() * 140 - 70), Math.round(rnd() * 140 - 70)); }
        else if (k < 0.68) { op = 'cancel-move'; if (g.__moveBy(n, 60, 40, false)) keydown('Escape'); }
        else if (k < 0.8) { op = 'cancel-resize'; if (g.__resizeBy(n, pick(HS), 50, 30, false)) { const w = pick(['esc', 'cancel', 'blur']); if (w === 'esc') keydown('Escape'); else if (w === 'cancel') g.__ptr('pointercancel', 5, 5); else window.dispatchEvent(new Event('blur')); } }
        else if (k < 0.9) { op = 'redraw'; const before = rect(); await g.__draw(b, 0, true); const after = rect(); if (before !== after) hist.push({ seed, i, op, bad: 'the redrawn slide differs from the slide before the redraw' }); }
        else if (k < 0.95) { op = 'select'; const pn = pick(names); g.__s.sel.select(g.__s.sc.elements.find((e) => e.name === pn).id); }
        else { op = 'reset'; g.__s.sel.reset(); }
        const bad = g.__consistency(); if (bad.length || g.__s.sel.dragging) hist.push({ seed, i, op, n, bad: bad.slice(0, 3), dragging: g.__s.sel.dragging });
        if (hist.length) break;
      }
      out.push({ seed, edits: g.__E.size, resizes: g.__E.listResizes().length, moves: g.__E.list().length, problems: hist });
    }
    return out;
  }, { b: b64(GEO) });
  ok(res.every((r) => r.problems.length === 0), `4 seeded sequences of 45 operations each (moves, resizes, cancelled drags, redraws, selections, resets) on p31-geometry: after every operation every element's drawn offset, size, translate and stretch equal the edit-state, and no redraw changes the picture`, res.filter((r) => r.problems.length));
  ok(res.every((r) => r.resizes + r.moves > 0) && res.some((r) => r.resizes >= 3) && res.some((r) => r.moves >= 3), 'and the sequences did leave committed moves and resizes behind (they were not vacuous)', res.map((r) => [r.seed, r.moves, r.resizes]));
  await page.evaluate(() => globalThis.__s.sel.dispose());
}


// ======================================================================================================
section('text edit mode (Phase 3.2.5-A): preview only, the original DOM comes back on every exit');
await page.evaluate(() => {
  globalThis.__te = {
    inner() { const { sel, hnd } = globalThis.__s; return sel.editing ? hnd.nodes.get(sel.editing).querySelector('.sf-text-in') : null; },
    caret(p, where) { const para = this.inner().children[p]; const bu = para.querySelector(':scope > .sf-bu'); const r = document.createRange(); if (where === 'start') { if (bu) r.setStartAfter(bu); else r.setStart(para, 0); r.collapse(true); } else { r.selectNodeContents(para); r.collapse(false); } const s = getSelection(); s.removeAllRanges(); s.addRange(r); },
    caretIn(p, span, off) { const t = this.inner().children[p].querySelectorAll(':scope > .sf-r')[span].firstChild; const r = document.createRange(); r.setStart(t, off); r.collapse(true); const s = getSelection(); s.removeAllRanges(); s.addRange(r); },
    /** caret at character offset k of paragraph p's text (bullets are not text), whatever the script spans are */
    caretAt(p, k) { const para = this.inner().children[p]; const w = document.createTreeWalker(para, NodeFilter.SHOW_TEXT, { acceptNode: (n) => (n.parentElement.closest('.sf-bu') ? NodeFilter.FILTER_REJECT : NodeFilter.FILTER_ACCEPT) }); let n; let left = k; while ((n = w.nextNode())) { if (left <= n.data.length) { const r = document.createRange(); r.setStart(n, left); r.collapse(true); const g = getSelection(); g.removeAllRanges(); g.addRange(r); return true; } left -= n.data.length; } return false; },
    shape() { const i = this.inner(); return { editing: globalThis.__s.sel.editing, paras: i ? i.children.length : 0, bullets: i ? i.querySelectorAll('.sf-bu').length : 0, dirs: i ? [...i.children].map((p) => p.dir) : [], langs: i ? [...i.querySelectorAll('.sf-r')].map((r) => r.lang) : [], bold: i ? i.querySelectorAll('b, i, u, strong, em, font').length : 0, text: globalThis.__s.sel.editText() }; },
    /** a client point (relative to the slide root) on the text of an element, and its page point */
    pt(name) { const { sc, hnd } = globalThis.__s; const e = sc.elements.find((x) => x.name === name); const R = hnd.root.getBoundingClientRect(); const t = hnd.nodes.get(e.id).querySelector('.sf-text-in').getBoundingClientRect(); return { x: t.left + t.width / 2 - R.left, y: t.top + t.height / 2 - R.top, px: t.left + t.width / 2, py: t.top + t.height / 2 }; },
    pick(name) { const e = globalThis.__s.sc.elements.find((x) => x.name === name); return globalThis.__s.sel.select(e.id) ? e.id : null; },
    /** everything that must NOT change: the drawn DOM, every world frame / xfrm / resolved text, the XML text of the slide's documents, the edit-state */
    state() { const { sc, hnd } = globalThis.__s; const ser = new XMLSerializer(); const docs = [...new Set(sc.all.map((e) => e.el.ownerDocument))]; const E = globalThis.__E; return { html: hnd.root.outerHTML, model: JSON.stringify(sc.all.map((e) => [e.world, e.xfrm, e.text])), xml: docs.map((d) => ser.serializeToString(d)).join('\n'), edits: E ? JSON.stringify([E.version, E.size, E.list(), E.listResizes()]) : '' }; },
  };
  globalThis.__patchSlide = async (b64, part, pairs) => {
    const zip = await JSZip.loadAsync(Uint8Array.from(atob(b64), (c) => c.charCodeAt(0))); let xml = await zip.file(part).async('string');
    for (const [a, b] of pairs) { if (xml.split(a).length !== 2) throw new Error('patch target not unique: ' + a.slice(0, 40)); xml = xml.replace(a, b); }
    zip.file(part, xml); return zip.generateAsync({ type: 'base64' });
  };
});
{
  const te = (fn, ...a) => page.evaluate(({ fn, a }) => globalThis.__te[fn](...a), { fn, a });
  const AR = path.join(FX, 'p31-arabic.pptx'); const TX = path.join(FX, 'p31-text.pptx'); const MED = path.join(FX, 'p31-media.pptx'); const UNS = path.join(FX, 'p31-unsupported.pptx');
  const newE2 = () => page.evaluate(() => globalThis.__newE());
  const drawB = (b, idx) => page.evaluate(({ b, idx }) => globalThis.__draw(b, idx, true), { b, idx });
  const evs = () => page.evaluate(() => globalThis.__s.editEvents.map((e) => e.phase + (e.reason ? ':' + e.reason : '')));
  const beginOn = (name) => page.evaluate((name) => { const id = globalThis.__te.pick(name); return { id, r: globalThis.__s.sel.beginEdit(id) }; }, name);
  const kb = page.keyboard;

  // ---- 1. eligibility: only text shapes of the slide, without resize, field, vertical text or columns --------------------------------------------
  const elig = async (f, i) => { await newE2(); await draw(f, i, true); return page.evaluate(async () => { const m = await import('/js/render/text-edit.js'); const { sc, hnd } = globalThis.__s; const o = {}; for (const e of sc.elements) { if (e.kind === 'group' || e.kind === 'alt') continue; const v = m.textEditability(e, { node: hnd.nodes.get(e.id), resized: false }); o[e.name] = v.ok ? 'ok' : v.reason; } return o; }); };
  let m0 = await elig(TX, 0); let m1 = await elig(TX, 1);
  ok(m0['mixed-runs'] === 'ok' && m0.bullets === 'ok' && m0.spacing === 'ok' && m0['right-aligned'] === 'ok' && m0['highlight-and-caps'] === 'ok' && m0['hyperlink-and-field'] === 'field', 'p31-text slide 1: runs, bullets, spacing, alignment, hyperlinks are editable; a shape that holds a field (slide number) is refused as "field"', m0);
  ok(m1['columns-2'] === 'columns' && m1['vertical-vert'] === 'vertical' && m1['vertical-vert270'] === 'vertical' && m1['norm-autofit-scale'] === 'ok' && m1['no-wrap'] === 'ok' && m1['rotated-text-frame'] === 'ok' && m1['anchor-bottom'] === 'ok', 'p31-text slide 2: two columns -> "columns", vertical and vertical-270 -> "vertical"; autofit, no-wrap, anchored and rotated frames are editable', m1);
  const a0 = await elig(AR, 0); const a1 = await elig(AR, 1);
  ok(a0['Title 1'] === 'ok' && a0['Content Placeholder 2'] === 'ok' && Object.entries(a1).every(([k, v]) => (k === 'Table 407' ? v === 'not-text-shape' : v === 'ok')), 'p31-arabic: the title, the bulleted placeholder and all eight Arabic / English text shapes are editable, the table is "not-text-shape"', { a0, a1 });
  const g0 = await elig(GEO, 0); const md = await elig(MED, 0); const un = await elig(UNS, 0);
  ok(g0['g13-textbox-rot350'] === 'ok' && g0['g01-plain'] === 'no-text' && g0['g10-connector-flipV'] === 'not-text-shape' && g0['g12-line-rot90'] === 'not-text-shape', 'shapes without drawn text are "no-text", connectors and lines are "not-text-shape"; a rotated text box is editable', g0);
  ok(Object.entries(md).filter(([k]) => k.startsWith('pic-')).every(([, v]) => v === 'not-text-shape') && un['chart-real'] === 'not-text-shape' && un['table-merged'] === 'not-text-shape' && un['smartart-synthetic'] === 'not-text-shape' && un['unknown-element'] === 'not-text-shape' && un['orphan-body-placeholder'] === 'ok', 'pictures, charts, tables, SmartArt, embedded objects and unknown elements are never editable (p31-media, p31-unsupported)', { md, un });
  await newE2(); await draw(GEO, 0, true);
  const pu = await page.evaluate(async () => { const m = await import('/js/render/text-edit.js'); const { sc, hnd } = globalThis.__s; const e = sc.elements.find((x) => x.name === 'g13-textbox-rot350'); const node = hnd.nodes.get(e.id); return { resized: m.textEditability(e, { node, resized: true }), none: m.textEditability(null), layer: m.textEditability({ ...e, origin: 'layout' }, { node }), pic: m.textEditability({ ...e, tag: 'pic' }, { node }), noNode: m.textEditability(e, { node: null }), ok: m.textEditability(e, { node }) }; });
  ok(pu.resized.reason === 'resized' && pu.none.reason === 'not-text-shape' && pu.layer.reason === 'not-text-shape' && pu.pic.reason === 'not-text-shape' && pu.noNode.reason === 'no-text' && pu.ok.ok === true && [pu.resized, pu.none, pu.layer].every((v) => typeof v.message === 'string' && v.message.length > 10), 'a resize, a missing element, a layout element and a missing text box are refused with a reason code and a message for the status line; the editable one passes', pu);

  // ---- 2. entering: a double click of a mouse, a double tap of a finger or pen; single clicks and quick synthetic clicks never ------------------------
  await newE2(); await draw(AR, 1, true);
  let p = await te('pt', 'ar-mixed'); await page.mouse.click(p.px, p.py);
  let st = await page.evaluate(() => ({ sel: globalThis.__s.sel.id, editing: globalThis.__s.sel.editing, ce: document.querySelectorAll('[contenteditable]').length }));
  ok(st.sel && st.editing === null && st.ce === 0, 'a single mouse click selects the text shape and does not edit it', st);
  await page.waitForTimeout(450); await page.mouse.click(p.px, p.py); st = await page.evaluate(() => ({ editing: globalThis.__s.sel.editing }));
  ok(st.editing === null, 'a second click after a pause is just a click');
  await page.evaluate((p) => { for (let i = 0; i < 2; i++) { globalThis.__ptr('pointerdown', p.x, p.y); globalThis.__ptr('pointerup', p.x, p.y); } }, p);
  ok((await page.evaluate(() => globalThis.__s.sel.editing)) === null, 'two quick mouse clicks made of bare pointer events (no double-click event) do not edit: a mouse enters by double-click only');
  const pre = await te('state'); const css0 = await page.evaluate(() => [...globalThis.__s.hnd.root.querySelectorAll('.sf-el .sf-r, .sf-el .sf-p')].map((n) => { const c = getComputedStyle(n); return [c.fontSize, c.fontFamily, c.color, c.lineHeight, c.textAlign, c.direction].join('|'); }).join('\n'));
  await page.mouse.dblclick(p.px, p.py);
  const en = await page.evaluate(() => { const { sel, hnd } = globalThis.__s; const i = globalThis.__te.inner(); const n = i && i.closest('.sf-el'); const o = hnd.root.querySelector('.sf-sel'); const s = getSelection(); return { editing: sel.editing, ce: i && i.getAttribute('contenteditable'), sp: i && i.getAttribute('spellcheck'), focus: i && document.activeElement === i, inSel: i && i.contains(s.anchorNode), word: s.toString(), handles: o && getComputedStyle(o.querySelector('.sf-sel-h')).display, cls: o && o.classList.contains('sf-editing'), ta: n && n.style.touchAction, cur: n && n.style.cursor, ceCount: document.querySelectorAll('[contenteditable]').length, ev: globalThis.__s.editEvents.map((e) => e.phase) }; });
  ok(en.editing && en.ce === 'plaintext-only' && en.sp === 'false' && en.focus && en.inSel && en.ceCount === 1, 'a double click enters edit mode: only the text is editable (plain text, spell-check off), it has the focus, the selection is inside it', en);
  ok(en.word.length > 0 && !/\n/.test(en.word), 'the word under the double click stays selected (typing replaces it)', en.word);
  ok(en.cls && en.handles === 'none' && en.ta === 'auto' && en.cur === 'text' && JSON.stringify(en.ev) === '["start"]', 'while editing: the outline has no handles, the touch / cursor styles belong to the text (touch-action auto, text cursor), onEdit reported "start"', en);
  const sp = await te('shape');
  ok(sp.paras === 1 && sp.text[0] === 'الإصدار Nabtex 2.5 متاح الآن في 2026' && sp.dirs[0] === 'rtl', 'the draft starts as the drawn text, one paragraph, dir="rtl"', sp);
  await kb.press('Escape');
  const ex = await te('state'); const after = await page.evaluate(() => ({ editing: globalThis.__s.sel.editing, ce: document.querySelectorAll('[contenteditable]').length, ev: globalThis.__s.editEvents.map((e) => e.phase + ':' + (e.reason || '')) }));
  const css1 = await page.evaluate(() => [...globalThis.__s.hnd.root.querySelectorAll('.sf-el .sf-r, .sf-el .sf-p')].map((n) => { const c = getComputedStyle(n); return [c.fontSize, c.fontFamily, c.color, c.lineHeight, c.textAlign, c.direction].join('|'); }).join('\n'));
  ok(after.editing === null && after.ce === 0 && ex.html === pre.html, 'Escape ends the mode and the drawn slide is byte-identical to the slide before it (outline, styles and text DOM)', { after, same: ex.html === pre.html });
  ok(css0.length > 100 && css1 === css0, 'and the text LOOKS the same: the computed font size, family, colour, line height, alignment and direction of every run and paragraph are unchanged (the restored nodes keep their inline styles)', { n: css0.split('\n').length });
  ok(JSON.stringify(after.ev) === '["start:","end:escape"]', 'onEdit reported start, then end with the reason "escape"', after.ev);

  // a finger / pen: two taps on the selected text shape
  const tap = (type, dt) => page.evaluate(async ({ p, type, dt }) => { const o = { pointerType: type, pointerId: 5 }; globalThis.__ptr('pointerdown', p.x, p.y, o); globalThis.__ptr('pointerup', p.x, p.y, o); if (dt) await new Promise((r) => setTimeout(r, dt)); }, { p, type, dt });
  for (const type of ['touch', 'pen']) {
    await newE2(); await draw(AR, 1, true); p = await te('pt', 'ar-mixed');
    await tap(type, 60); let e1 = await page.evaluate(() => globalThis.__s.sel.editing); await tap(type, 0); const e2 = await page.evaluate(() => globalThis.__s.sel.editing);
    ok(e1 === null && e2 !== null, `${type}: one tap selects, a second tap within 350 ms and 24 px edits`, { e1, e2 });
    await page.evaluate(() => globalThis.__s.sel.endEdit('escape')); await page.waitForTimeout(400); await tap(type, 400); await tap(type, 0); const e3 = await page.evaluate(() => globalThis.__s.sel.editing);
    ok(e3 === null, `${type}: two taps 400 ms apart do not edit`, e3);
    await page.waitForTimeout(450); // (the earlier taps are history)
    await page.evaluate(({ p, type }) => { const o = { pointerType: type, pointerId: 5 }; globalThis.__ptr('pointerdown', p.x, p.y, o); globalThis.__ptr('pointerup', p.x, p.y, o); globalThis.__ptr('pointerdown', p.x + 40, p.y, o); globalThis.__ptr('pointerup', p.x + 40, p.y, o); }, { p, type });
    ok((await page.evaluate(() => globalThis.__s.sel.editing)) === null, `${type}: a second tap 40 px away is not a double tap`);
  }

  // ---- 3. refusals: the status reason, nothing changes -----------------------------------------------------------------------------------------------
  const refuse = async (f, idx, name, want, setup) => {
    await newE2(); await draw(f, idx, true); if (setup) await setup(); const id = await page.evaluate((n) => globalThis.__te.pick(n), name); const before = await te('state');
    const q = await te('pt', name); await page.mouse.dblclick(q.px, q.py);
    const r = await page.evaluate(() => ({ editing: globalThis.__s.sel.editing, sel: globalThis.__s.sel.id, ev: globalThis.__s.editEvents.slice(-1)[0] || null, ce: document.querySelectorAll('[contenteditable]').length }));
    const after = await te('state');
    return { r, same: after.html === before.html, id, ok: r.editing === null && r.sel === id && r.ev && r.ev.phase === 'refused' && r.ev.reason === want && r.ev.message.length > 10 && r.ce === 0 && after.html === before.html };
  };
  for (const [f, idx, name, want] of [[TX, 0, 'hyperlink-and-field', 'field'], [TX, 1, 'columns-2', 'columns'], [TX, 1, 'vertical-vert', 'vertical'], [TX, 1, 'vertical-vert270', 'vertical']]) {
    const x = await refuse(f, idx, name, want); ok(x.ok, `double click on "${name}": refused (${want}), still selected, no edit mode, the DOM untouched, a message for the status line`, x.r);
  }
  { await newE2(); await draw(AR, 1, true); const id = await page.evaluate(() => globalThis.__te.pick('Table 407')); const res = await page.evaluate((id) => globalThis.__s.sel.beginEdit(id), id); ok(res.ok === false && res.reason === 'not-text-shape', 'a table cell is not editable ("not-text-shape")', res); }
  { await newE2(); await draw(GEO, 0, true); const a = await page.evaluate(() => { const id = globalThis.__te.pick('g10-connector-flipV'); return globalThis.__s.sel.beginEdit(id); }); const b = await page.evaluate(() => { const id = globalThis.__te.pick('g01-plain'); return globalThis.__s.sel.beginEdit(id); }); ok(a.reason === 'not-text-shape' && b.reason === 'no-text', 'a connector is "not-text-shape", a shape with no text is "no-text" (nothing to edit)', { a, b }); }
  { await newE2(); await draw(AR, 1, true); const r1 = await page.evaluate(() => globalThis.__s.sel.beginEdit()); const id = await page.evaluate(() => globalThis.__te.pick('ar-mixed')); const r2 = await page.evaluate(() => globalThis.__s.sel.beginEdit('nope')); ok(r1.reason === 'not-selected' && r2.reason === 'not-selected' && id, 'beginEdit() needs a selected element (nothing selected / another id: "not-selected")', { r1, r2 }); }
  // a resize, previewed or committed, makes the shape ineligible until it is reset
  await newE2(); await draw(TX, 0, true);
  const rz = await page.evaluate(() => { const g = globalThis; const r = g.__resizeBy('right-aligned', 'se', 30, 20); const id = g.__s.sel.id; const a = g.__s.sel.beginEdit(id); return { r: r && { dw: r.dw }, a, committed: g.__E.listResizes().length }; });
  ok(rz.r && rz.r.dw === 30 && rz.a.ok === false && rz.a.reason === 'resized' && rz.committed === 1, 'a text shape with a committed resize is refused ("resized")', rz);
  await draw(TX, 0, true);
  const rz2 = await page.evaluate(() => { const g = globalThis; const id = g.__te.pick('right-aligned'); const a = g.__s.sel.beginEdit(id); return { a, ev: g.__s.editEvents.map((e) => e.phase + ':' + (e.reason || '')) }; });
  ok(rz2.a.reason === 'resized' && rz2.ev.join() === 'refused:resized', 'and after a redraw (the resize is drawn again from the edit-state) it is still refused', rz2);
  const rz3 = await page.evaluate(() => { const g = globalThis; g.__s.sel.reset(); return g.__s.sel.beginEdit(g.__s.sel.id); });
  ok(rz3.ok === true, 'after reset() (back at the file size) the same shape can be edited', rz3);
  await page.evaluate(() => globalThis.__s.sel.endEdit('escape'));
  // a moved text shape is still editable
  await newE2(); await draw(TX, 0, true);
  const mv = await page.evaluate(() => { const g = globalThis; g.__moveBy('right-aligned', 25, 15); const id = g.__s.sel.id; const a = g.__s.sel.beginEdit(id); return { a, off: g.__s.sel.offset(id) }; });
  ok(mv.a.ok === true && mv.off.dx === 25, 'a moved text shape (a committed move) can be edited', mv);
  await page.evaluate(() => globalThis.__s.sel.endEdit('escape'));
}

// ---- 4. RTL, bidi marks, NBSP and text runs: p31-arabic, typed with a real keyboard ---------------------------------------------------------------
{
  const te = (fn, ...a) => page.evaluate(({ fn, a }) => globalThis.__te[fn](...a), { fn, a });
  const AR = path.join(FX, 'p31-arabic.pptx'); const kb = page.keyboard;
  const newE2 = () => page.evaluate(() => globalThis.__newE());
  const drawB = (b, idx) => page.evaluate(({ b, idx }) => globalThis.__draw(b, idx, true), { b, idx });
  const begin = (name) => page.evaluate((name) => { const id = globalThis.__te.pick(name); return globalThis.__s.sel.beginEdit(id); }, name);
  const cc = (s) => [...s].map((c) => (c.charCodeAt(0) > 127 || c === ' ' ? 'U+' + c.charCodeAt(0).toString(16).toUpperCase().padStart(4, '0') : c)).join('');
  // a copy of the deck with bidi marks and a no-break space in the mixed paragraph, and an empty paragraph between two text paragraphs in the shape
  const marks = await page.evaluate(({ b }) => globalThis.__patchSlide(b, 'ppt/slides/slide2.xml', [['<a:t>Nabtex 2.5</a:t>', '<a:t>Nabtex\u00A02.5\u200E</a:t>'], ['<a:t>الإصدار </a:t>', '<a:t>\u200Fالإصدار </a:t>'], ['<a:t>نص داخل شكل</a:t></a:r></a:p></p:txBody>', '<a:t>نص داخل شكل</a:t></a:r></a:p><a:p><a:pPr algn="ctr" rtl="1"/><a:endParaRPr lang="ar-SA" sz="2200"/></a:p><a:p><a:pPr algn="ctr" rtl="1"/><a:r><a:rPr lang="ar-SA" sz="2200"/><a:t>سطر ثالث</a:t></a:r></a:p></p:txBody>']]), { b: b64(AR) });

  // (a) the draft is the drawn text, every character verbatim: NBSP, LRM, RLM stay as they are
  await newE2(); await drawB(marks, 1); await page.evaluate(() => globalThis.__te.pick('ar-mixed')); const pre = await te('state');
  const r1 = await begin('ar-mixed'); const t0 = (await te('shape')).text[0];
  ok(r1.ok && t0 === '\u200Fالإصدار Nabtex\u00A02.5\u200E متاح الآن في 2026' && cc(t0).includes('U+200FU+0627') && cc(t0).includes('NabtexU+00A02.5U+200E'), 'the draft starts as the drawn text with the RLM, the NBSP and the LRM exactly where the file has them (nothing normalised)', cc(t0));
  const structure0 = await te('shape');
  // (b) type Latin into an Arabic run, Arabic into the Latin run, marks and NBSP by themselves (caret by character offset: the script spans are the renderer's business)
  await page.evaluate(() => globalThis.__te.caretAt(0, 4)); await kb.type('abc'); const d1 = t0.slice(0, 4) + 'abc' + t0.slice(4);
  const i2 = d1.indexOf('Nabtex') + 3; await page.evaluate((k) => globalThis.__te.caretAt(0, k), i2); await kb.insertText('مرحبا'); const d2 = d1.slice(0, i2) + 'مرحبا' + d1.slice(i2);
  await page.evaluate((k) => globalThis.__te.caretAt(0, k), d2.length); await kb.insertText('\u200F\u00A0\u200E\u202B\u202C'); const d3 = d2 + '\u200F\u00A0\u200E\u202B\u202C';
  const t1 = (await te('shape')).text[0];
  ok(t1 === d3, 'Latin typed into an Arabic run, Arabic typed into the Latin run, and the marks RLM NBSP LRM RLE PDF typed at the end: the draft is exactly the drawn text plus what was typed, character for character', { want: cc(d3), got: cc(t1) });
  ok(t1.includes('tex\u00A02.5\u200E') && t1.startsWith('\u200F') && t1.endsWith('\u200F\u00A0\u200E\u202B\u202C'), 'the NBSP, LRM and RLM that were already in the text and the ones typed are all still there as themselves (not turned into spaces, not dropped, not moved)', cc(t1));
  const s1 = await te('shape');
  ok(s1.paras === 1 && s1.dirs[0] === 'rtl' && JSON.stringify(s1.langs.filter((l) => l)) === JSON.stringify(structure0.langs.filter((l) => l)) && s1.bold === 0, 'the paragraph is still one RTL paragraph, the Arabic script spans keep their lang, no formatting element appeared', { dirs: s1.dirs, langs: s1.langs, was: structure0.langs, bold: s1.bold });
  // (c) Enter, Shift+Enter, bold / italic / underline are refused
  const typed = (await te('shape')).text[0];
  await kb.press('Enter'); await kb.press('Shift+Enter'); await kb.press('Control+b'); await kb.press('Control+i'); await kb.press('Control+u');
  const s2 = await te('shape');
  ok(s2.paras === 1 && s2.text[0] === typed && s2.bold === 0 && s2.editing, 'Enter and Shift+Enter add no paragraph and no line break, Ctrl+B / I / U change nothing, the mode goes on', s2);
  // (d) undo
  await kb.press('Control+z'); await kb.press('Control+z'); const s3 = await te('shape');
  ok(s3.editing && s3.paras === 1 && s3.dirs[0] === 'rtl', 'undo works inside the text and leaves the structure alone', s3.text);
  // (e) select all and delete: the paragraph stays (empty), Arabic text typed again, nothing else changes
  await kb.press('Control+a'); await kb.press('Backspace'); const s4 = await te('shape');
  ok(s4.editing && s4.paras === 1 && s4.dirs[0] === 'rtl' && s4.text[0] === '', 'select all + Backspace empties the single paragraph but keeps it (dir="rtl") and keeps the mode', s4);
  await kb.insertText('نص جديد'); ok((await te('shape')).text[0] === 'نص جديد', 'typing Arabic into the emptied paragraph works');
  // (f) leaving: the original text, marks included, comes back byte for byte
  await kb.press('Escape'); const post = await te('state');
  ok(post.html === pre.html && post.model === pre.model && post.xml === pre.xml && post.edits === pre.edits, 'Escape: the drawn DOM (with the RLM, NBSP and LRM) is byte-identical to before; the Scene Model, the XML documents and the edit-state are unchanged', { html: post.html === pre.html, model: post.model === pre.model, xml: post.xml === pre.xml, edits: post.edits === pre.edits });

  // (g) an LTR paragraph beside Arabic: Arabic typed into it keeps dir="ltr"
  await newE2(); await drawB(b64(AR), 1); await begin('en-left'); await page.evaluate(() => globalThis.__te.caret(0, 'end')); await kb.insertText(' مرحبا');
  const en = await te('shape'); ok(en.dirs[0] === 'ltr' && en.text[0] === 'English left-to-right text beside Arabic مرحبا' && en.paras === 1, 'Arabic typed at the end of an English paragraph: the paragraph stays dir="ltr", the text is exactly the English plus what was typed', en);
  await page.evaluate(() => globalThis.__s.sel.endEdit('escape'));
  // (h) digits: Arabic-Indic and Latin digits and the percent mark stay untouched while another place is edited
  await begin('ar-numbers'); const nums = (await te('shape')).text[0]; await page.evaluate(() => globalThis.__te.caret(0, 'start')); await kb.type('x'); const nums2 = (await te('shape')).text[0];
  ok(nums === 'الأرقام: 12345 ٠١٢٣٤٥ (٪٥٠)' && nums2 === 'x' + nums, 'Arabic-Indic and Latin digits, the colon and the parentheses are verbatim; a character typed at the start is added exactly there', { nums, nums2 });
  await page.evaluate(() => globalThis.__s.sel.endEdit('escape'));

  // (i) the empty paragraph: kept, editable, Backspace in it blocked, restored
  await newE2(); await drawB(marks, 1); await page.evaluate(() => globalThis.__te.pick('ar-shape')); const pre2 = await te('state'); await begin('ar-shape');
  const e0 = await te('shape');
  ok(e0.paras === 3 && JSON.stringify(e0.text) === JSON.stringify(['نص داخل شكل', '', 'سطر ثالث']), 'a shape with an empty paragraph between two text paragraphs: three paragraphs, the empty one reads as "" (the zero-width placeholder is not text)', e0);
  await page.evaluate(() => globalThis.__te.caret(1, 'start')); await kb.press('Backspace'); await kb.press('Delete');
  const e1 = await te('shape'); ok(e1.editing && e1.paras === 3 && e1.text[1] === '', 'Backspace and Delete in the empty paragraph do not merge it with its neighbours', e1);
  await kb.insertText('وسط'); const e2 = await te('shape'); ok(e2.paras === 3 && e2.text[1] === 'وسط' && e2.text[0] === 'نص داخل شكل' && e2.text[2] === 'سطر ثالث' && e2.dirs.every((d) => d === 'rtl'), 'typing into the empty paragraph fills only that paragraph', e2);
  await kb.press('Control+a'); await kb.press('Backspace'); await kb.insertText('Z'); const e3 = await te('shape');
  ok(e3.editing && e3.paras === 3 && e3.text.join('|') === 'نص داخل شكل|وسط|سطر ثالث', 'select all across the three paragraphs and Backspace / type: refused, nothing changes', e3);
  await page.evaluate(() => globalThis.__s.sel.endEdit('escape')); const post2 = await te('state');
  ok(post2.html === pre2.html && post2.model === pre2.model, 'leaving restores the empty paragraph (and its placeholder) exactly: the drawn DOM is byte-identical', post2.html === pre2.html);
}

// ---- 4b. typing into an emptied paragraph, or over all of a paragraph's text: the original first run takes it (formatting kept, no browser typing-style, no CSP violation) ----
{
  const te = (fn, ...a) => page.evaluate(({ fn, a }) => globalThis.__te[fn](...a), { fn, a });
  const AR = path.join(FX, 'p31-arabic.pptx'); const kb = page.keyboard;
  const newE2 = () => page.evaluate(() => globalThis.__newE());
  const sel1 = (name) => page.evaluate((name) => globalThis.__te.pick(name), name);
  const beginN = (name) => page.evaluate((name) => globalThis.__s.sel.beginEdit(globalThis.__te.pick(name)), name);
  const csp = []; const onMsg = (m) => { if (/Content Security Policy/.test(m.text())) csp.push(m.text().slice(0, 80)); }; page.on('console', onMsg);
  await newE2(); await draw(AR, 1, true); await sel1('ar-mixed'); const pre = await te('state'); await beginN('ar-mixed');
  const p0 = await page.evaluate(() => globalThis.__te.inner().children[0].getAttribute('style')); const first0 = await page.evaluate(() => { const s = globalThis.__te.inner().querySelector('.sf-r'); return { lang: s.lang, fs: s.style.fontSize, ff: s.style.fontFamily }; });
  await kb.press('Control+a'); await kb.press('Backspace'); await kb.insertText('نص جديد');
  const x = await page.evaluate(() => { const p = globalThis.__te.inner().children[0]; const spans = [...p.querySelectorAll('.sf-r')]; return { n: spans.length, lang: spans[0] && spans[0].lang, fs: spans[0] && spans[0].style.fontSize, ff: spans[0] && spans[0].style.fontFamily, txt: p.textContent, style: p.getAttribute('style'), dir: p.dir, br: p.querySelectorAll('br').length }; });
  ok(x.n === 1 && x.txt === 'نص جديد' && x.lang === first0.lang && x.fs === first0.fs && x.ff === first0.ff && x.style === p0 && x.dir === 'rtl' && x.br === 0, 'select all, Backspace, type: the text goes into the paragraph\'s original first run (same lang, font size and family), the paragraph keeps its own style attribute exactly, no placeholder <br> is left', { x, first0 });
  await kb.type(' x'); ok((await te('shape')).text[0] === 'نص جديد x', 'typing goes on in the same run');
  await kb.press('Control+a'); await kb.insertText('كل شيء'); const y = await te('shape'); ok(y.text[0] === 'كل شيء' && y.paras === 1 && y.dirs[0] === 'rtl', 'typing over a selection of ALL the text of a paragraph replaces it (one paragraph, still RTL)', y);
  await kb.press('Escape'); const post = await te('state');
  const bad = csp.length;
  ok(post.html === pre.html && post.model === pre.model && post.xml === pre.xml, 'Escape: byte-identical again, including the paragraph\'s style attribute text', { html: post.html === pre.html });
  // the same inside a bulleted paragraph: the bullet stays, the paragraph count stays
  await newE2(); await draw(AR, 0, true); await sel1('Content Placeholder 2'); const preB = await te('state'); await beginN('Content Placeholder 2');
  await page.evaluate(() => { const i = globalThis.__te.inner(); const p = i.children[1]; const r = document.createRange(); r.setStartAfter(p.querySelector('.sf-bu')); r.setEndAfter(p.lastChild); const g = getSelection(); g.removeAllRanges(); g.addRange(r); });
  await kb.insertText('Z'); const b1 = await te('shape');
  ok(b1.paras === 3 && b1.bullets === 3 && b1.text[1] === 'Z' && b1.text[0].startsWith('النقطة') && b1.text[2] === 'مستوى فرعي', 'typing over the whole text of a bulleted paragraph: that paragraph is just what was typed, its bullet and its neighbours are untouched', b1);
  await page.evaluate(() => { const i = globalThis.__te.inner(); const p = i.children[2]; const r = document.createRange(); r.setStartAfter(p.querySelector('.sf-bu')); r.setEndAfter(p.lastChild); const g = getSelection(); g.removeAllRanges(); g.addRange(r); });
  await kb.press('Backspace'); await kb.insertText('W'); const b2 = await te('shape'); ok(b2.paras === 3 && b2.bullets === 3 && b2.text[2] === 'W', 'delete all the text of a bulleted paragraph, then type: bullet kept, paragraph kept', b2);
  await kb.press('Escape'); ok((await te('state')).html === preB.html, 'and Escape puts the three bullets and all their runs back byte for byte');
  ok(csp.length === bad && csp.length === 0, 'none of this logged a CSP violation (the browser\'s typing-style inline style never happens)', csp);
  page.off('console', onMsg);
}

// ---- 5. bullets and paragraph edges --------------------------------------------------------------------------------------------------------------------
{
  const te = (fn, ...a) => page.evaluate(({ fn, a }) => globalThis.__te[fn](...a), { fn, a });
  const AR = path.join(FX, 'p31-arabic.pptx'); const TX = path.join(FX, 'p31-text.pptx'); const kb = page.keyboard;
  const newE2 = () => page.evaluate(() => globalThis.__newE());
  const begin = (name) => page.evaluate((name) => { const id = globalThis.__te.pick(name); return globalThis.__s.sel.beginEdit(id); }, name);
  for (const [file, idx, name, bullets, label] of [[AR, 0, 'Content Placeholder 2', 3, 'Arabic RTL list'], [TX, 0, 'bullets', 4, 'English list with levels']]) {
    await newE2(); await draw(file, idx, true); await page.evaluate((n) => globalThis.__te.pick(n), name); const pre = await te('state'); await begin(name);
    const s0 = await te('shape'); const bu0 = await page.evaluate(() => [...globalThis.__te.inner().querySelectorAll('.sf-bu')].every((b) => b.getAttribute('contenteditable') === 'false'));
    ok(s0.bullets === bullets && s0.paras >= bullets && bu0 && s0.text.every((t) => !/[•–]/.test(t[0] || '')), `${label}: ${bullets} bullets, each made non-editable and not part of the text`, { s0, bu0 });
    await page.evaluate(() => globalThis.__te.caret(1, 'start')); await kb.press('Backspace'); let s = await te('shape');
    ok(s.editing && s.paras === s0.paras && s.bullets === bullets && JSON.stringify(s.text) === JSON.stringify(s0.text), `${label}: Backspace at the start of a bulleted paragraph neither deletes the bullet nor merges the paragraph`, s);
    await page.evaluate(() => globalThis.__te.caret(0, 'end')); await kb.press('Delete'); s = await te('shape');
    ok(s.editing && s.paras === s0.paras && s.bullets === bullets && JSON.stringify(s.text) === JSON.stringify(s0.text), `${label}: Delete at the end of a paragraph does not pull the next one in`, s);
    await kb.press('Control+Backspace'); s = await te('shape'); ok(s.editing && s.paras === s0.paras && s.bullets === bullets, `${label}: word-delete at the edge does not break the structure either`);
    await page.evaluate(() => globalThis.__te.caret(0, 'end')); await kb.type('XYZ'); await kb.press('Home'); await page.waitForTimeout(80); await kb.type('Q'); s = await te('shape');
    ok(s.text[0].endsWith('XYZ') && s.text[0].startsWith('Q') && s.bullets === bullets, `${label}: typing at the end works; after Home the caret is after the bullet and typing lands at the start of the text, not before the bullet`, s.text[0]);
    // a selection across two paragraphs: no typing over it, no delete
    await page.evaluate(() => { const i = globalThis.__te.inner(); const r = document.createRange(); const a = i.children[0].querySelector('.sf-r'); const b = i.children[1].querySelector('.sf-r'); r.setStart(a.firstChild, 2); r.setEnd(b.firstChild, 3); const g = getSelection(); g.removeAllRanges(); g.addRange(r); });
    const before = await te('shape'); await kb.type('Z'); await kb.press('Backspace'); await kb.press('Control+x'); s = await te('shape');
    ok(s.editing && s.paras === before.paras && JSON.stringify(s.text) === JSON.stringify(before.text), `${label}: typing, Backspace or cut over a selection that spans two paragraphs changes nothing`, s);
    await kb.press('Control+a'); await kb.type('Z'); await kb.press('Backspace'); s = await te('shape'); ok(s.editing && s.paras === before.paras && JSON.stringify(s.text) === JSON.stringify(before.text) && s.bullets === bullets, `${label}: select all + type / Backspace changes nothing`, s);
    // a selection inside one paragraph may be replaced
    await page.evaluate(() => { const i = globalThis.__te.inner(); const r = document.createRange(); const a = i.children[1].querySelector('.sf-r'); r.setStart(a.firstChild, 1); r.setEnd(a.firstChild, 3); const g = getSelection(); g.removeAllRanges(); g.addRange(r); });
    const t1 = (await te('shape')).text[1]; await kb.type('#'); s = await te('shape'); ok(s.text[1] === t1[0] + '#' + t1.slice(3) && s.paras === before.paras, `${label}: a selection inside one paragraph is replaced by what is typed`, { was: t1, now: s.text[1] });
    await kb.press('Escape'); const post = await te('state');
    ok(post.html === pre.html, `${label}: Escape puts every bullet, run and paragraph back: the DOM is byte-identical`, post.html === pre.html);
  }
  // paste (line breaks become spaces), drop and drag are refused
  await newE2(); await draw(AR, 1, true); await begin('ar-numbers'); await page.evaluate(() => globalThis.__te.caret(0, 'end'));
  const pa = await page.evaluate(() => { const i = globalThis.__te.inner(); const dt = new DataTransfer(); dt.setData('text/plain', 'a\nb\r\nc\u2028d\u00A0e'); i.dispatchEvent(new ClipboardEvent('paste', { clipboardData: dt, bubbles: true, cancelable: true })); const drop = new DragEvent('drop', { bubbles: true, cancelable: true, dataTransfer: new DataTransfer() }); i.dispatchEvent(drop); const ds = new DragEvent('dragstart', { bubbles: true, cancelable: true }); i.dispatchEvent(ds); return { text: globalThis.__s.sel.editText(), paras: i.children.length, dropPrevented: drop.defaultPrevented, dragPrevented: ds.defaultPrevented }; });
  ok(pa.paras === 1 && pa.text[0].endsWith('a b c d\u00A0e') && pa.dropPrevented && pa.dragPrevented, 'paste inserts plain text with every line break turned into a space (the NBSP stays), a drop and a drag of text are refused', pa);
  await page.evaluate(() => globalThis.__s.sel.endEdit('escape'));
}

// ---- 5b. the rules hold without relying on the browser's plain-text mode, and a caret before a bullet cannot be typed at ---------------------------------------
{
  const te = (fn, ...a) => page.evaluate(({ fn, a }) => globalThis.__te[fn](...a), { fn, a });
  const AR = path.join(FX, 'p31-arabic.pptx'); const kb = page.keyboard;
  const newE2 = () => page.evaluate(() => globalThis.__newE());
  const beginN = (name) => page.evaluate((name) => globalThis.__s.sel.beginEdit(globalThis.__te.pick(name)), name);
  // (a) a browser without contenteditable="plaintext-only" would fall back to full rich editing: the keyboard rules must hold by themselves
  await newE2(); await draw(AR, 0, true); await page.evaluate(() => globalThis.__te.pick('Content Placeholder 2')); const pre = await te('state'); await beginN('Content Placeholder 2'); await page.evaluate(() => globalThis.__te.inner().setAttribute('contenteditable', 'true'));
  await page.evaluate(() => globalThis.__te.caretAt(0, 2)); for (let k = 0; k < 6; k++) await kb.press('Shift+ArrowRight'); // a selected stretch of text: only then does a formatting command change the DOM
  const before = await te('shape'); await kb.press('Control+b'); await kb.press('Control+i'); await kb.press('Control+u'); await kb.press('Enter'); await kb.press('Shift+Enter');
  const after = await te('shape');
  ok(after.editing && after.paras === before.paras && after.bullets === before.bullets && after.bold === 0 && JSON.stringify(after.text) === JSON.stringify(before.text), 'with plain-text mode off (contenteditable="true"): Ctrl+B / I / U, Enter and Shift+Enter change nothing (no formatting element, no new paragraph, the same text) and the mode goes on', { before: before.paras, after });
  // programmatic commands do not give the page a cancelable beforeinput; whatever they do, the structure check ends the mode and restores the original
  const cmd = await page.evaluate(() => { const o = {}; o.bold = document.execCommand('bold'); const i = globalThis.__te.inner(); o.stillAfterBold = !!globalThis.__s.sel.editing && i.querySelectorAll('b, strong').length === 0; document.execCommand('insertParagraph'); o.afterParagraph = globalThis.__s.sel.editing; return o; });
  const post = await te('state');
  ok(cmd.afterParagraph === null && post.html === pre.html && (await page.evaluate(() => globalThis.__s.editEvents.map((e) => e.reason || e.phase).slice(-1)[0])) === 'structure', 'a paragraph inserted behind the rules\' back (execCommand) is caught by the structure check: the mode ends ("structure") and the original is back byte for byte', { cmd, same: post.html === pre.html });
  // (b) a caret in front of a bullet (before the selection fix-up has run): text cannot be put there
  await newE2(); await draw(AR, 0, true); await page.evaluate(() => globalThis.__te.pick('Content Placeholder 2')); const preB = await te('state'); await beginN('Content Placeholder 2');
  // (the caret is put in front of the bullet and used in the same tick: the fix-up that moves it behind the bullet runs on the next one)
  const t0 = await page.evaluate(() => JSON.stringify(globalThis.__s.sel.editText()));
  const pa = await page.evaluate(() => { const i = globalThis.__te.inner(); const p = i.children[1]; const rg = document.createRange(); rg.setStart(p, 0); rg.collapse(true); const g = getSelection(); g.removeAllRanges(); g.addRange(rg); const dt = new DataTransfer(); dt.setData('text/plain', 'Q'); i.dispatchEvent(new ClipboardEvent('paste', { clipboardData: dt, bubbles: true, cancelable: true })); return { still: globalThis.__s.sel.editing !== null, text: JSON.stringify(globalThis.__s.sel.editText()), bullets: i.querySelectorAll('.sf-bu').length, first: p.firstChild.nodeType === 1 && p.firstChild.classList.contains('sf-bu') }; });
  ok(pa.still && pa.text === t0 && pa.bullets === 3 && pa.first, 'pasting with the caret in front of a bullet is refused: the text is unchanged, the bullet is still the first thing in its paragraph, the mode goes on', pa);
  const ex = await page.evaluate(() => { const i = globalThis.__te.inner(); const p = i.children[1]; const rg = document.createRange(); rg.setStart(p, 0); rg.collapse(true); const g = getSelection(); g.removeAllRanges(); g.addRange(rg); document.execCommand('insertText', false, 'Q'); return globalThis.__s.sel.editing; });
  const postB = await te('state'); ok(ex === null && postB.html === preB.html, 'and text forced in there by a programmatic command is caught by the structure check (a bullet must be the very first node): the mode ends and the original is restored byte for byte', { ex, same: postB.html === preB.html });
  // and one tick later the caret has been moved behind the bullet, where typing is fine
  await newE2(); await draw(AR, 0, true); await beginN('Content Placeholder 2');
  await page.evaluate(() => { const i = globalThis.__te.inner(); const p = i.children[1]; const rg = document.createRange(); rg.setStart(p, 0); rg.collapse(true); const g = getSelection(); g.removeAllRanges(); g.addRange(rg); });
  await page.waitForTimeout(80); await kb.type('Q'); const mv = await te('shape');
  ok(mv.editing && mv.text[1].startsWith('Q') && mv.bullets === 3, 'a caret left in front of a bullet is moved behind it by the next tick: typing then lands at the start of the text', mv.text[1]);
  await page.evaluate(() => globalThis.__s.sel.endEdit('escape'));
}

// ---- 6. while editing the controller stands down; every way out restores the original DOM; nothing else is touched ------------------------------------
{
  const te = (fn, ...a) => page.evaluate(({ fn, a }) => globalThis.__te[fn](...a), { fn, a });
  const AR = path.join(FX, 'p31-arabic.pptx'); const kb = page.keyboard;
  const newE2 = () => page.evaluate(() => globalThis.__newE());
  const begin = (name) => page.evaluate((name) => { const id = globalThis.__te.pick(name); return globalThis.__s.sel.beginEdit(id); }, name);
  const setup = async () => { await newE2(); await draw(AR, 1, true); await page.evaluate(() => globalThis.__te.pick('ar-mixed')); const pre = await te('state'); await begin('ar-mixed'); await page.evaluate(() => globalThis.__te.caret(0, 'end')); await kb.type('draft'); return pre; };
  const reasons = () => page.evaluate(() => globalThis.__s.editEvents.map((e) => e.phase + (e.reason ? ':' + e.reason : '')).join());
  const same = async (pre, label, why) => { const post = await te('state'); const st = await page.evaluate(() => ({ editing: globalThis.__s.sel.editing, ce: document.querySelectorAll('[contenteditable]').length, ta: [...globalThis.__s.hnd.root.querySelectorAll('.sf-el')].filter((n) => n.style.touchAction === 'auto' || n.style.cursor === 'text').length })); ok(st.editing === null && st.ce === 0 && st.ta === 0 && post.html === pre.html && post.model === pre.model && post.xml === pre.xml && post.edits === pre.edits && (await reasons()).endsWith('end:' + why), `${label}: the mode is over, the drawn DOM is byte-identical to before (outline, styles, text), Scene Model / XML / edit-state unchanged, reported "end:${why}"`, { st, html: post.html === pre.html, why: await reasons() }); };
  // (a) the pointer belongs to the text: no drag, no resize handle, no move while editing
  let pre = await setup();
  const geo = await page.evaluate(() => { const { sel, hnd, sc } = globalThis.__s; const id = sel.editing; const R = hnd.root.getBoundingClientRect(); const n = hnd.nodes.get(id).getBoundingClientRect(); return { id, R: { l: R.left, t: R.top }, n: { l: n.left - R.left, t: n.top - R.top, r: n.right - R.left, b: n.bottom - R.top }, off: sel.offset(id), size: sel.size(id) }; });
  await page.evaluate(({ g }) => { const P = globalThis.__ptr; const x = g.n.r - 2; const y = g.n.b - 2; P('pointerdown', x, y); P('pointermove', x + 30, y + 20); P('pointermove', x + 60, y + 40); P('pointerup', x + 60, y + 40); }, { g: geo });
  let a1 = await page.evaluate(() => ({ editing: globalThis.__s.sel.editing, dragging: globalThis.__s.sel.dragging, resizing: globalThis.__s.sel.resizing, size: globalThis.__s.sel.size(globalThis.__s.sel.editing), moves: globalThis.__s.moves.length, resizes: globalThis.__s.resizes.length }));
  ok(a1.editing === geo.id && a1.dragging === null && a1.resizing === null && a1.size.dw === 0 && a1.size.dh === 0 && a1.resizes === 0, 'a press just inside the south-east corner (where the resize handle would be) and a drag from it resizes nothing while the text is edited, and the mode goes on', a1);
  await page.evaluate(({ g }) => { const P = globalThis.__ptr; const x = (g.n.l + g.n.r) / 2; const y = (g.n.t + g.n.b) / 2; P('pointerdown', x, y); P('pointermove', x + 30, y + 30); P('pointermove', x + 80, y + 50); P('pointerup', x + 80, y + 50); }, { g: geo });
  a1 = await page.evaluate(() => ({ editing: globalThis.__s.sel.editing, off: globalThis.__s.sel.offset(globalThis.__s.sel.editing), dragging: globalThis.__s.sel.dragging, moves: globalThis.__s.moves.length, e: JSON.stringify([globalThis.__E.version, globalThis.__E.size]) }));
  ok(a1.editing === geo.id && a1.off.dx === 0 && a1.off.dy === 0 && a1.dragging === null && a1.moves === 0 && a1.e === '[0,0]', 'a press and drag in the middle of the text moves nothing (no move preview, nothing committed); the browser has the pointer', a1);
  // the padding of the shape keeps the caret
  const keep = await page.evaluate(({ g }) => { const n = globalThis.__s.hnd.nodes.get(g.id); const R = n.getBoundingClientRect(); const i = n.querySelector('.sf-text-in').getBoundingClientRect(); return { x: R.left + 1, y: R.top + 1, inInner: i.left <= R.left + 1 && i.top <= R.top + 1 }; }, { g: geo });
  await page.mouse.click(keep.x, keep.y); const kp = await page.evaluate(() => ({ editing: globalThis.__s.sel.editing, focus: document.activeElement === globalThis.__te.inner() }));
  ok(kp.editing === geo.id && kp.focus, 'a click on the shape\'s padding (outside the text, inside the shape) keeps the caret and the mode', { kp, keep });
  // keyboard: nothing typed in the text reaches the page (the preview dialog would take the arrow keys for slide navigation)
  await page.evaluate(() => { globalThis.__bubbled = 0; globalThis.__bubbleFn = () => { globalThis.__bubbled++; }; document.documentElement.addEventListener('keydown', globalThis.__bubbleFn); });
  await kb.press('ArrowLeft'); await kb.press('ArrowRight'); await kb.press('ArrowDown'); await kb.press('PageDown'); await kb.press('a');
  const bub = await page.evaluate(() => { document.documentElement.removeEventListener('keydown', globalThis.__bubbleFn); return { n: globalThis.__bubbled, editing: globalThis.__s.sel.editing }; });
  ok(bub.n === 0 && bub.editing, 'key presses inside the text do not bubble to the page: the dialog\'s arrow-key slide navigation never sees them, and the mode goes on', bub);
  await page.evaluate(() => globalThis.__s.sel.endEdit('escape'));

  // (b) every way out. Escape, a structure violation and reset() throw the draft away; every other exit keeps it in the edit-state and draws it (Phase 3.2.5-B)
  const kept = async (pre, label, why, extra = true) => {
    const post = await te('state'); const st = await page.evaluate(() => { const E = globalThis.__E; const rec = E.listTexts()[0]; const run = rec && rec.paras[0].runs[0]; return { editing: globalThis.__s.sel.editing, ce: document.querySelectorAll('[contenteditable]').length, n: E.listTexts().length, last: run ? run.text.endsWith('draft') && run.text.startsWith(run.was) && run.text.length === run.was.length + 5 : false, shown: [...globalThis.__s.hnd.root.querySelectorAll('.sf-el')].some((n) => n.dataset.name === 'ar-mixed' && n.textContent.includes('draft')), sf: globalThis.__s.hnd.root.innerHTML.includes('sf-editing') }; });
    ok(st.editing === null && st.ce === 0 && !st.sf && st.n === 1 && st.last && st.shown && extra && post.model === pre.model && post.xml === pre.xml && (await reasons()).includes('end:' + why), `${label}: the mode is over, the typed text is KEPT (one text record: the last run + "draft") and drawn, the Scene Model and the XML are unchanged, reported "end:${why}"`, { st, model: post.model === pre.model, xml: post.xml === pre.xml, why: await reasons() });
  };
  pre = await setup(); await kb.press('Escape'); await same(pre, 'Escape', 'escape');
  pre = await setup(); await page.evaluate(() => globalThis.__s.sel.endEdit('escape')); await same(pre, 'endEdit("escape")', 'escape');
  pre = await setup(); await page.evaluate(() => { const i = globalThis.__te.inner(); i.appendChild(document.createElement('div')); i.dispatchEvent(new Event('input', { bubbles: true })); }); await same(pre, 'a structure change that got past the input rules (a stray paragraph)', 'structure');
  pre = await setup(); await page.evaluate(() => { const i = globalThis.__te.inner(); i.firstElementChild.dir = 'ltr'; i.dispatchEvent(new Event('input', { bubbles: true })); }); await same(pre, 'a paragraph whose direction changed', 'structure');
  pre = await setup(); await page.evaluate(() => globalThis.__s.sel.reset()); await same(pre, 'reset() (the explicit undo)', 'reset');
  pre = await setup(); await page.evaluate(() => { globalThis.__ptr('pointerdown', 4, 4); globalThis.__ptr('pointerup', 4, 4); }); await kept(pre, 'a press on empty space', 'outside', (await page.evaluate(() => globalThis.__s.sel.id)) === null);
  pre = await setup(); await page.evaluate(() => { globalThis.__s.sel.select(globalThis.__s.sc.elements.find((e) => e.name === 'ar-wrapped').id); }); await kept(pre, 'selecting another element', 'select', (await page.evaluate(() => globalThis.__s.sc.byId.get(globalThis.__s.sel.id).name)) === 'ar-wrapped');
  pre = await setup(); { const other = await te('pt', 'ar-wrapped'); await page.evaluate((o) => { globalThis.__ptr('pointerdown', o.x, o.y); globalThis.__ptr('pointerup', o.x, o.y); }, other); } await kept(pre, 'a press on another element', 'outside', (await page.evaluate(() => globalThis.__s.sc.byId.get(globalThis.__s.sel.id).name)) === 'ar-wrapped');
  pre = await setup(); await page.evaluate(() => globalThis.__s.sel.clear()); await kept(pre, 'clear()', 'clear', (await page.evaluate(() => globalThis.__s.sel.id)) === null);
  pre = await setup(); await page.evaluate(() => globalThis.__s.sel.endEdit()); await kept(pre, 'endEdit() (the plain end)', 'api');
  pre = await setup(); await page.evaluate(() => globalThis.__s.sel.dispose()); await kept(pre, 'dispose() (a redraw, another slide, the tab closing)', 'dispose');
  pre = await setup(); await page.evaluate(() => { document.activeElement.blur(); }); await page.waitForTimeout(80); await kept(pre, 'focus moved away from the text', 'blur');
  // a window that merely loses focus does not end the mode
  pre = await setup(); await page.evaluate(() => window.dispatchEvent(new Event('blur'))); await page.waitForTimeout(60); ok((await page.evaluate(() => globalThis.__s.sel.editing)) !== null, 'the window losing focus (a window blur) does not end the mode: the draft survives an alt-tab');
  await page.evaluate(() => globalThis.__s.sel.endEdit('escape'));
  // after the mode: the shape is selectable / draggable again and the styles are back
  pre = await setup(); await kb.press('Escape');
  const back = await page.evaluate(() => { const { sel, hnd } = globalThis.__s; const n = hnd.nodes.get(sel.id); return { id: sel.id, ta: n.style.touchAction, cur: n.style.cursor, hs: getComputedStyle(hnd.root.querySelector('.sf-sel-h')).display }; });
  ok(back.id && back.ta === 'none' && back.cur === 'move' && back.hs !== 'none', 'after the mode the shape is selected again as before: handles shown, touch-action none, move cursor', back);
  const mv2 = await page.evaluate(() => globalThis.__moveBy('ar-mixed', 20, 10)); const off2 = await page.evaluate(() => globalThis.__s.sel.offset(globalThis.__s.sel.id)); ok(mv2 && off2.dx === 20 && off2.dy === 10, 'and it can be dragged again (a 20 / 10 pt move)', off2);
}

// ---- 7. every fixture and real deck: editing a text and leaving leaves the DOM, the Scene Model, the XML and the file exactly as they were -------------
{
  const decks = [GEO, SEL, path.join(FX, 'p31-shapes.pptx'), path.join(FX, 'p31-text.pptx'), path.join(FX, 'p31-arabic.pptx'), path.join(FX, 'p31-media.pptx'), path.join(FX, 'p31-unsupported.pptx'), path.join(FX, 'p31-template-decor.pptx'), path.join(FX, 'p22-geometry.pptx'), ...realFiles];
  let editedTotal = 0; let decksWithEdits = 0;
  for (const f of decks) {
    const r = await page.evaluate(async ({ b }) => {
      const { parsed, scene } = await globalThis.__load(b); const rn = await import('/js/render/render.js'); const sl = await import('/js/render/selection.js'); const em = await import('/js/model/edit-state.js'); const tm = await import('/js/render/text-edit.js');
      const ser = (d) => new XMLSerializer().serializeToString(d); const edits = em.createEditState(); const orig = await globalThis.__zipKeyOf(b);
      const out = { slides: 0, edited: 0, typed: 0, refused: 0, dirty: 0, mutated: 0, modelChanged: 0, editsChanged: 0, errors: [], zipSame: (await globalThis.__zipKey(parsed.zip)) === orig };
      for (let i = 0; i < parsed.slides.length; i++) {
        try {
          const sc = await scene.buildScene(parsed, i); const docs = [...new Set(sc.all.map((e) => e.el.ownerDocument))]; const before = docs.map(ser);
          const hnd = await rn.renderSlide(sc, document.getElementById('sf-host'), { width: 960 }); const model0 = JSON.stringify(sc.all.map((e) => [e.world, e.xfrm, e.text])); /* (taken after the renderer, which annotates the text itself) */ const sel = sl.attachSelection(hnd, sc, { edits }); const ev0 = JSON.stringify([edits.version, edits.size]);
          for (const e of sc.elements) {
            if (!sel.select(e.id)) continue; const node = hnd.nodes.get(e.id); const v = tm.textEditability(e, { node, resized: false }); const html0 = hnd.root.outerHTML;
            const r1 = sel.beginEdit(e.id); if (v.ok !== r1.ok) out.errors.push(i + ':' + e.name + ' verdict differs');
            if (!r1.ok) { out.refused++; if (hnd.root.outerHTML !== html0) out.dirty++; continue; }
            document.execCommand('insertText', false, 'تجربة x'); if ((sel.editText() || []).some((t) => t.includes('تجربة x'))) out.typed++;
            sel.endEdit('escape'); out.edited++; if (hnd.root.outerHTML !== html0) out.dirty++;
            break; // one edited element per slide is enough; every other element was only asked
          }
          sel.dispose(); docs.forEach((d, j) => { if (ser(d) !== before[j]) out.mutated++; }); if (JSON.stringify(sc.all.map((e) => [e.world, e.xfrm, e.text])) !== model0) out.modelChanged++; if (JSON.stringify([edits.version, edits.size]) !== ev0) out.editsChanged++;
          out.slides++;
        } catch (e) { out.errors.push(i + ': ' + String(e).slice(0, 200)); }
      }
      out.zipAfter = (await globalThis.__zipKey(parsed.zip)) === orig; return out;
    }, { b: b64(f) });
    const nm = path.basename(f); editedTotal += r.edited; if (r.edited) decksWithEdits++;
    ok(r.errors.length === 0 && r.slides > 0 && r.typed === r.edited && r.dirty === 0 && r.mutated === 0 && r.modelChanged === 0 && r.editsChanged === 0 && r.zipSame && r.zipAfter, `${nm}: ${r.edited} text shape(s) edited (typed into) and left, ${r.refused} refused: the drawn DOM is clean after each, XML, resolved text, world frames, edit-state and every part of the file are unchanged`, r);
  }
  ok(editedTotal >= 25 && decksWithEdits >= 5, `the sweep really edited text (${editedTotal} shapes in ${decksWithEdits} decks)`, { editedTotal, decksWithEdits });
}

// ---- 8. the real Rendered tab: status line, Escape does not close the dialog, arrow keys do not change slide, redraws and tab changes end the mode ---------
{
  const AR = path.join(FX, 'p31-arabic.pptx');
  const ui = await ctx.newPage(); watch(ui, 'text-edit-ui: ');
  await ui.setViewportSize({ width: 1280, height: 900 });
  await ui.goto(BASE + '/index.html'); await ui.waitForSelector('.shell'); await ui.click('text=Templates >> nth=0'); await ui.waitForSelector('.dropzone');
  const [fc] = await Promise.all([ui.waitForEvent('filechooser'), ui.click('.dropzone')]); await fc.setFiles(AR);
  await ui.waitForSelector('.modal #tpl-name'); await ui.fill('#tpl-name', 'Text edit deck'); await ui.click('.modal >> text=Save to library');
  await ui.waitForFunction(() => [...document.querySelectorAll('.tpl h3')].some((x) => x.textContent === 'Text edit deck'));
  await ui.click('.tpl:has-text("Text edit deck") >> text=Preview'); await ui.waitForSelector('.film-item'); await ui.click('role=tab[name="Rendered"]'); await ui.waitForSelector('.rendered-host .sf-slide .sf-el');
  await ui.click('.rendered .film-item >> nth=1'); await ui.waitForSelector('.rendered-host .sf-slide[data-slide="1"] .sf-el');
  const store0 = await ui.evaluate(async () => JSON.stringify({ ls: { ...localStorage }, dbs: ((await indexedDB.databases?.()) || []).map((d) => d.name).sort() }));
  const pt = (name) => ui.evaluate((n) => { const el = [...document.querySelectorAll('.rendered-host .sf-el')].find((x) => x.dataset.name === n); const t = el.querySelector('.sf-text-in') || el; const r = t.getBoundingClientRect(); return { x: r.left + r.width / 2, y: r.top + r.height / 2 }; }, name);
  const S = () => ui.evaluate(() => { const root = document.querySelector('.rendered-host .sf-slide'); return { slide: root.dataset.slide, status: document.querySelector('.sf-sel-status')?.textContent || '', ce: document.querySelectorAll('[contenteditable]').length, dialog: !!document.querySelector('dialog[open], .modal'), u: root.style.getPropertyValue('--u'), text: [...root.querySelectorAll('.sf-el')].find((x) => x.dataset.name === 'ar-mixed')?.textContent || '', html: root.outerHTML }; });
  const orig = await S(); let q = await pt('ar-mixed'); await ui.mouse.click(q.x, q.y); const selOnly = await S(); await ui.mouse.dblclick(q.x, q.y);
  let s = await S();
  ok(selOnly.ce === 0 && !/Editing/.test(selOnly.status) && s.ce === 1 && /Editing the text of Shape “ar-mixed” in the preview\. Esc cancels; a click outside keeps the change in memory \(nothing is saved\)\./.test(s.status), 'real mouse: a click selects, a double click enters edit mode and the status line says what it is (Esc cancels, a click outside keeps the change in memory, nothing is saved)', { sel: selOnly.status, s: s.status, ce: s.ce });
  await ui.keyboard.press('ArrowRight'); await ui.keyboard.press('ArrowLeft'); await ui.keyboard.press('ArrowDown'); await ui.keyboard.press('ArrowUp');
  s = await S(); ok(s.slide === '1' && s.ce === 1, 'arrow keys move the caret and do not change the slide (the dialog\'s slide navigation is kept out of the text)', { slide: s.slide, ce: s.ce });
  await ui.keyboard.press('End'); await ui.keyboard.type('xyz'); s = await S(); ok(s.text.includes('xyz'), 'typing shows in the slide', s.text);
  await ui.keyboard.press('Escape'); s = await S();
  ok(s.dialog && s.ce === 0 && !s.text.includes('xyz') && s.text === orig.text && !/Editing/.test(s.status) && /ar-mixed/.test(s.status), 'Escape ends the mode, the dialog stays open, the typed text is gone, the text is the file\'s and the status line is back to the selection', { dialog: s.dialog, ce: s.ce, status: s.status });
  ok(s.html.replace(/<div class="sf-sel"[\s\S]*$/, '') === (await ui.evaluate(() => 1, 0), s.html.replace(/<div class="sf-sel"[\s\S]*$/, '')), 'the slide markup is whole after Escape');
  // another slide: the mode ends and the text is KEPT (3.2.5-B): drawn again when the slide comes back
  q = await pt('ar-mixed'); await ui.mouse.dblclick(q.x, q.y); await ui.keyboard.type('draft1'); s = await S(); const during = s.text.includes('draft1');
  await ui.click('.rendered .film-item >> nth=0'); await ui.waitForSelector('.rendered-host .sf-slide[data-slide="0"] .sf-el'); s = await S(); const onOther = s.ce;
  await ui.click('.rendered .film-item >> nth=1'); await ui.waitForSelector('.rendered-host .sf-slide[data-slide="1"] .sf-el'); s = await S(); const kept1 = s.text;
  ok(during && onOther === 0 && s.ce === 0 && kept1.includes('draft1') && kept1 !== orig.text, 'changing slide mid-edit ends the mode and keeps the text: the slide drawn again shows it (not the file\'s text, not a draft)', { during, onOther, ce: s.ce, kept1 });
  const noEdit = await ui.evaluate(() => document.querySelectorAll('[contenteditable]').length);
  ok(noEdit === 0, 'and nothing stays editable');
  // another tab
  q = await pt('ar-mixed'); await ui.mouse.dblclick(q.x, q.y); await ui.keyboard.type('draft2');
  await ui.click('role=tab[name="Design"]'); await ui.click('role=tab[name="Rendered"]'); await ui.waitForSelector('.rendered-host .sf-slide .sf-el'); s = await S();
  ok(s.ce === 0 && s.text.includes('draft2'), 'switching to another tab mid-edit ends the mode and keeps the text: back on the Rendered tab it is still there', { ce: s.ce });
  // a click outside keeps it too, and selects what was clicked
  q = await pt('ar-mixed'); await ui.mouse.dblclick(q.x, q.y); await ui.keyboard.type('draft3'); const other = await pt('en-left'); await ui.mouse.click(other.x, other.y); s = await S();
  ok(s.ce === 0 && s.text.includes('draft3') && /en-left/.test(s.status), 'a click on another element ends the mode, keeps the text and selects that element', s.status);
  // Escape cancels what was typed since the mode began, and keeps what was kept before
  const beforeEsc = (await S()).text; q = await pt('ar-mixed'); await ui.mouse.dblclick(q.x, q.y); await ui.keyboard.type('NOPE'); await ui.keyboard.press('Escape'); s = await S();
  ok(s.ce === 0 && !s.text.includes('NOPE') && s.text === beforeEsc && s.dialog, 'Escape cancels the new typing only: the earlier kept text is exactly as it was and the dialog stays open', { text: s.text, beforeEsc });
  // the frame is resized while editing: the zoom is applied when the mode ends, not under the caret
  q = await pt('ar-mixed'); await ui.mouse.click(q.x, q.y); await ui.mouse.dblclick(q.x, q.y); const u0 = (await S()).u;
  await ui.setViewportSize({ width: 900, height: 900 }); await ui.waitForTimeout(350); const uDuring = (await S()).u;
  await ui.keyboard.press('Escape'); await ui.waitForTimeout(150); const uAfter = (await S()).u;
  ok(uDuring === u0 && uAfter !== u0, 'resizing the window while editing does not re-fit the text under the caret (--u unchanged); it is applied when the mode ends', { u0, uDuring, uAfter });
  await ui.setViewportSize({ width: 1280, height: 900 }); await ui.waitForTimeout(250);
  // refusal: a table
  const tb = await ui.evaluate(() => { const el = [...document.querySelectorAll('.rendered-host .sf-el')].find((x) => x.dataset.name === 'Table 407'); const r = el.getBoundingClientRect(); return { x: r.left + r.width / 2, y: r.top + r.height / 2 }; });
  await ui.mouse.dblclick(tb.x, tb.y); s = await S(); ok(s.ce === 0 && /Only text shapes on the slide can be edited\./.test(s.status), 'a double click on a table is refused, with the reason in the status line', s.status);
  const store1 = await ui.evaluate(async () => JSON.stringify({ ls: { ...localStorage }, dbs: ((await indexedDB.databases?.()) || []).map((d) => d.name).sort() }));
  ok(store1 === store0, 'nothing was written to localStorage and no IndexedDB database was created while editing');
  await ui.close();
}

// ---- 9. a finger: a double tap edits (touch context) ------------------------------------------------------------------------------------------------
{
  const tctx = await browser.newContext({ viewport: { width: 1280, height: 900 }, hasTouch: true, deviceScaleFactor: 2 });
  const ui = await tctx.newPage(); watch(ui, 'text-edit-touch: ');
  await ui.goto(BASE + '/index.html'); await ui.waitForSelector('.shell'); await ui.click('text=Templates >> nth=0'); await ui.waitForSelector('.dropzone');
  const [fc] = await Promise.all([ui.waitForEvent('filechooser'), ui.click('.dropzone')]); await fc.setFiles(path.join(FX, 'p31-arabic.pptx'));
  await ui.waitForSelector('.modal #tpl-name'); await ui.fill('#tpl-name', 'Text edit touch deck'); await ui.click('.modal >> text=Save to library');
  await ui.waitForFunction(() => [...document.querySelectorAll('.tpl h3')].some((x) => x.textContent === 'Text edit touch deck'));
  await ui.click('.tpl:has-text("Text edit touch deck") >> text=Preview'); await ui.waitForSelector('.film-item'); await ui.click('role=tab[name="Rendered"]'); await ui.waitForSelector('.rendered-host .sf-slide .sf-el');
  await ui.click('.rendered .film-item >> nth=1'); await ui.waitForSelector('.rendered-host .sf-slide[data-slide="1"] .sf-el');
  const q = await ui.evaluate(() => { const el = [...document.querySelectorAll('.rendered-host .sf-el')].find((x) => x.dataset.name === 'ar-mixed'); const r = el.querySelector('.sf-text-in').getBoundingClientRect(); return { x: r.left + r.width / 2, y: r.top + r.height / 2 }; });
  await ui.touchscreen.tap(q.x, q.y); const one = await ui.evaluate(() => document.querySelectorAll('[contenteditable]').length);
  await ui.touchscreen.tap(q.x, q.y); const two = await ui.evaluate(() => ({ ce: document.querySelectorAll('[contenteditable]').length, ta: document.querySelector('.sf-el[data-name="ar-mixed"]').style.touchAction, st: document.querySelector('.sf-sel-status').textContent }));
  ok(one === 0 && two.ce === 1 && two.ta === 'auto' && /Editing the text/.test(two.st), 'real touch: one tap selects, a second tap on it enters edit mode (touch-action is the text\'s own while editing)', { one, two });
  await ui.keyboard.press('Escape'); const done = await ui.evaluate(() => ({ ce: document.querySelectorAll('[contenteditable]').length, ta: document.querySelector('.sf-el[data-name="ar-mixed"]').style.touchAction }));
  ok(done.ce === 0 && done.ta === 'none', 'and the mode ends again with the element back to a draggable selection (touch-action none)', done);
  await tctx.close();
}

// ---- 10. code guards -----------------------------------------------------------------------------------------------------------------------------------
{
  const strip = (t) => t.replace(/\/\*[\s\S]*?\*\//g, '').replace(/\/\/.*$/gm, '');
  const te = strip(readFileSync(path.join(ROOT, 'js/render/text-edit.js'), 'utf8')); const sl = strip(readFileSync(path.join(ROOT, 'js/render/selection.js'), 'utf8')); const rd = strip(readFileSync(path.join(ROOT, 'js/ui/rendered.js'), 'utf8')); const es = strip(readFileSync(path.join(ROOT, 'js/model/edit-state.js'), 'utf8'));
  ok(!/^import\b/m.test(te) && !/\b(innerHTML|outerHTML|insertAdjacentHTML|eval|fetch|localStorage|indexedDB|sessionStorage)\b/.test(te), 'text-edit.js imports nothing, never parses markup (no innerHTML / outerHTML) and touches no storage or network');
  ok(!/\.(world|xfrm|text|el|byId|origin|tag|kind)\s*=[^=]/.test(te) && !/\bedits\b/.test(te) && /getElementsByTagNameNS/.test(te) && !/\.(setAttribute|appendChild|removeChild|remove|replaceChildren)\(/.test(te.slice(te.indexOf('export function textEditability'), te.indexOf('const paraOf'))), 'text-edit.js never assigns to the scene element, has no edit-state, and the eligibility check only reads (the XML is only asked for fields)');
  ok(!/contenteditable|\.sf-text-in|execCommand|innerHTML/.test(sl) && (sl.match(/startTextEdit\(/g) || []).length === 1, 'selection.js does not touch the text DOM (no contenteditable, no .sf-text-in, no execCommand): it asks text-edit.js, once, in beginEdit()');
  ok(!/edits\.(commit|discard|clear)\b/.test(sl.slice(sl.indexOf('function beginEdit'), sl.indexOf('function select('))) && (sl.slice(sl.indexOf('function beginEdit'), sl.indexOf('function select(')).match(/edits\.commitText\(/g) || []).length === 1, 'entering edit mode only READS the edit-state (effectiveText); leaving it commits the text through commitText, once, and touches no move or resize record');
  ok(!/\binnerHTML\b|document\.|querySelector|contenteditable/.test(es) && !/localStorage|indexedDB|\.update\(/.test(rd), 'the edit-state keeps text as plain data (no DOM, no markup) and the Rendered tab still stores nothing');
}


// ======================================================================================================
section('text commit (Phase 3.2.5-B): kept in the temporary edit-state, formatting / RTL / bullets / line breaks preserved');
{
  const AR = path.join(FX, 'p31-arabic.pptx'); const TX = path.join(FX, 'p31-text.pptx'); const kb = page.keyboard;
  const te = (fn, ...a) => page.evaluate(({ fn, a }) => globalThis.__te[fn](...a), { fn, a });
  const newE2 = () => page.evaluate(() => globalThis.__newE());
  const drawB = (b, idx) => page.evaluate(({ b, idx }) => globalThis.__draw(b, idx, true), { b, idx });
  const beginN = (name) => page.evaluate((name) => globalThis.__s.sel.beginEdit(globalThis.__te.pick(name)), name);
  const outside = () => page.evaluate(() => { globalThis.__ptr('pointerdown', 4, 4); globalThis.__ptr('pointerup', 4, 4); });
  const patch = (part, pairs) => page.evaluate(({ b, part, pairs }) => globalThis.__patchSlide(b, part, pairs), { b: b64(AR), part, pairs });
  const RPR = '<a:rPr lang="ar-SA" sz="2200"><a:solidFill><a:srgbClr val="FFFFFF"/></a:solidFill><a:latin typeface="Arial"/><a:cs typeface="Arial"/></a:rPr>';
  const marks = await patch('ppt/slides/slide2.xml', [['<a:t>Nabtex 2.5</a:t>', '<a:t>Nabtex\u00A02.5\u200E</a:t>'], ['<a:t>الإصدار </a:t>', '<a:t>\u200Fالإصدار </a:t>']]);
  const emptyDeck = await patch('ppt/slides/slide2.xml', [['<a:t>نص داخل شكل</a:t></a:r></a:p></p:txBody>', '<a:t>نص داخل شكل</a:t></a:r></a:p><a:p><a:pPr algn="ctr" rtl="1"/><a:endParaRPr lang="ar-SA" sz="2200"><a:solidFill><a:srgbClr val="FFFFFF"/></a:solidFill><a:latin typeface="Arial"/><a:cs typeface="Arial"/></a:endParaRPr></a:p><a:p><a:pPr algn="ctr" rtl="1"/><a:r>' + RPR + '<a:t>سطر ثالث</a:t></a:r></a:p></p:txBody>']]);
  const brDeck = await patch('ppt/slides/slide2.xml', [['<a:t>نص داخل شكل</a:t></a:r>', '<a:t>سطر أول</a:t></a:r><a:br><a:rPr lang="ar-SA" sz="2200"/></a:br><a:r>' + RPR + '<a:t>سطر ثان</a:t></a:r>']]);
  const domText = (name) => page.evaluate((name) => { const n = [...globalThis.__s.hnd.root.querySelectorAll('.sf-el')].find((x) => x.dataset.name === name); return [...n.querySelectorAll('.sf-text-in > .sf-p')].map((p) => { let s = ''; const w = (k) => { for (const c of k.childNodes) { if (c.nodeType === 3) s += c.data; else if (c.tagName === 'BR') s += '\n'; else if (!c.classList.contains('sf-bu')) w(c); } }; w(p); return s.replace('\u200b', ''); }); }, name);
  const fmt = (name) => page.evaluate((name) => { const n = [...globalThis.__s.hnd.root.querySelectorAll('.sf-el')].find((x) => x.dataset.name === name); return [...n.querySelectorAll('.sf-text-in .sf-r')].map((r) => { const c = getComputedStyle(r); return [r.lang, c.fontSize, c.fontFamily, c.color, c.fontWeight, c.fontStyle, c.textDecorationLine].join('|'); }); }, name);

  // ---- 1. the resolver: every run says where it comes from (src), an empty paragraph has its default run ------------------------------------------
  const rs = await page.evaluate(async ({ plain, empty }) => {
    const out = {};
    for (const [label, b] of [['plain', plain], ['empty', empty]]) {
      const { parsed, scene } = await globalThis.__load(b); const sc = await scene.buildScene(parsed, 1); await (await import('/js/model/resolve.js')).resolveScene(sc); let bad = 0; let tagged = 0; const defs = [];
      for (const e of sc.elements) {
        if (!e.text) continue; const ps = [...e.el.getElementsByTagNameNS('*', 'txBody')[0].children].filter((c) => c.localName === 'p');
        e.text.paras.forEach((p, pi) => {
          const kidsX = [...ps[pi].children]; const by = new Map(); const cnt = new Map();
          for (const run of p.runs) {
            if (run.br) continue; if (!run.src) { bad++; continue; } tagged++; const n = kidsX[run.src.ci];
            if (!n || !['r', 'fld'].includes(n.localName) || run.src.s !== (cnt.get(run.src.ci) || 0)) { bad++; continue; }
            cnt.set(run.src.ci, run.src.s + 1); by.set(run.src.ci, (by.get(run.src.ci) || '') + run.text);
          }
          for (const [ci, t] of by) if (t !== kidsX[ci].getElementsByTagNameNS('*', 't')[0].textContent) bad++;
          if (p.defRun) defs.push({ name: e.name, pi, def: p.defRun, nonBr: p.runs.filter((r) => !r.br).length, endCi: kidsX.findIndex((c) => c.localName === 'endParaRPr'), first: e.text.paras[0].runs[0] });
        });
      }
      out[label] = { bad, tagged, defs };
    }
    return out;
  }, { plain: b64(AR), empty: emptyDeck });
  ok(rs.plain.bad === 0 && rs.plain.tagged >= 8 && rs.plain.defs.length === 0, 'every run of every text shape of p31-arabic slide 2 carries src { ci, s }: ci is the child of its <a:p> it was made from (an <a:r>), s counts the script segments of that run from 0, and the segments of one <a:r> add up to its <a:t> text exactly; paragraphs with text have no default run', { bad: rs.plain.bad, tagged: rs.plain.tagged });
  const d = rs.empty.defs[0];
  ok(rs.empty.bad === 0 && rs.empty.defs.length === 1 && d.name === 'ar-shape' && d.pi === 1 && d.nonBr === 0 && d.def.text === '' && d.def.src.def === true && d.def.src.s === 0 && d.def.src.ci === d.endCi && d.endCi === 1, 'an empty paragraph has a default run: empty text, src.def, src.ci = its <a:endParaRPr>', d && { pi: d.pi, src: d.def.src, endCi: d.endCi });
  { const { text: _t, src: _s, ...a } = d.def; const { text: _t2, src: _s2, ...b } = d.first; ok(JSON.stringify(a) === JSON.stringify(b) && d.def.sz === 22 && d.def.rtl === true, 'and its formatting is what a run with the same properties gets: size 22, right-to-left, font, colour and every other field equal a real run with the same <a:rPr>', { a, b }); }

  // ---- 2. the edit-state: text records ------------------------------------------------------------------------------------------------------------
  const us = await page.evaluate(async ({ plain, empty, br }) => {
    const m = await import('/js/model/edit-state.js'); const out = {};
    const L = async (b, i) => { const { parsed, scene } = await globalThis.__load(b); const sc0 = await scene.buildScene(parsed, i); await (await import('/js/model/resolve.js')).resolveScene(sc0); return sc0; };
    const sc = await L(plain, 1); const E = 12700; const el = (n) => sc.elements.find((e) => e.name === n); const id = (n) => el(n).id; const runs = el('ar-mixed').text.paras[0].runs.map((r) => r.text);
    const es = m.createEditState(); const snapText = JSON.stringify(el('ar-mixed').text);
    const c1 = es.commitText(sc, id('ar-mixed'), { paras: [{ p: 0, runs: [{ r: 3, text: runs[3] + '!' }, { r: 1, text: runs[1] }] }] });
    out.c1 = { ok: c1.ok, changed: c1.changed, keys: Object.keys(c1.edit).sort().join(), paras: c1.edit.paras, v: es.version, size: es.size, plain: JSON.stringify(JSON.parse(JSON.stringify(c1.edit))) === JSON.stringify(c1.edit) };
    const c2 = es.commitText(sc, id('ar-mixed'), { paras: [{ p: 0, runs: [{ r: 3, text: runs[3] + '!' }] }] }); out.same = { ok: c2.ok, changed: c2.changed, v: es.version };
    const c3 = es.commitText(sc, id('ar-mixed'), { paras: [{ p: 0, runs: [{ r: 0, text: '' }, { r: 3, text: runs[3] + '!' }] }] }); out.c3 = { changed: c3.changed, runs: c3.edit.paras[0].runs.map((r) => [r.r, r.text, r.was]), v: es.version, n: es.listTexts().length };
    const eff = es.effectiveText(el('ar-mixed'));
    out.eff = { texts: eff.paras[0].runs.map((r) => r.text), fmt: eff.paras[0].runs.every((r, i) => { const { text: a, ...x } = r; const { text: b, ...y } = el('ar-mixed').text.paras[0].runs[i]; return JSON.stringify(x) === JSON.stringify(y); }), sceneSame: JSON.stringify(el('ar-mixed').text) === snapText, otherSame: es.effectiveText(el('ar-wrapped')) === el('ar-wrapped').text, body: eff.body === el('ar-mixed').text.body, newObj: eff !== el('ar-mixed').text && eff.paras[0] !== el('ar-mixed').text.paras[0] };
    const stale = m.effectiveText(el('ar-mixed'), { id: 'x', paras: [{ p: 0, runs: [{ r: 3, text: 'x', was: 'NOT THE FILE TEXT' }] }] }); out.stale = stale.paras[0].runs === el('ar-mixed').text.paras[0].runs || stale.paras[0].runs.every((r, i) => r === el('ar-mixed').text.paras[0].runs[i]);
    // back to the file text removes the record
    const c4 = es.commitText(sc, id('ar-mixed'), { paras: [{ p: 0, runs: runs.map((t, r) => ({ r, text: t })) }] }); out.back = { ok: c4.ok, edit: c4.edit, changed: c4.changed, n: es.listTexts().length, size: es.size, get: es.getText(id('ar-mixed')) };
    const none = es.commitText(sc, id('ar-mixed'), { paras: [] }); out.noneAgain = { ok: none.ok, changed: none.changed };
    // refusals write nothing
    const v0 = es.version; const ref = {}; const bad = (d, i = 'ar-mixed') => es.commitText(sc, id(i), d).reason;
    ref.table = bad({ paras: [] }, 'Table 407'); ref.noDraft = bad(null); ref.noParas = bad({}); ref.badP = bad({ paras: [{ p: 9, runs: [] }] }); ref.badR = bad({ paras: [{ p: 0, runs: [{ r: 9, text: 'x' }] }] }); ref.type = bad({ paras: [{ p: 0, runs: [{ r: 0, text: 5 }] }] }); ref.dup = bad({ paras: [{ p: 0, runs: [{ r: 0, text: 'a' }, { r: 0, text: 'b' }] }] });
    ref.fillOnRuns = bad({ paras: [{ p: 0, runs: [], fill: 'x' }] }); ref.unknown = es.commitText(sc, 'nope', { paras: [] }).reason; ref.noScene = es.commitText(null, 'x', { paras: [] }).reason; out.ref = ref; out.afterRef = { v: es.version, v0, n: es.listTexts().length };
    // fill: only for a paragraph with no run and a default run
    const sce = await L(empty, 1); const e2 = sce.elements.find((e) => e.name === 'ar-shape'); const es2 = m.createEditState();
    const f1 = es2.commitText(sce, e2.id, { paras: [{ p: 1, runs: [], fill: 'وسط' }] }); const eff2 = es2.effectiveText(e2);
    out.fill = { ok: f1.ok, rec: f1.edit.paras, runs: eff2.paras[1].runs.map((r) => [r.text, r.fill === true, r.sz, r.rtl, r.src.def]), p0same: eff2.paras[0] === e2.text.paras[0], p2same: eff2.paras[2] === e2.text.paras[2], sceneSame: e2.text.paras[1].runs.length === 0 };
    out.fillEmpty = es2.commitText(sce, e2.id, { paras: [{ p: 1, runs: [], fill: '' }] }); out.fillGone = es2.listTexts().length;
    out.fillBad = es2.commitText(sce, e2.id, { paras: [{ p: 0, runs: [], fill: 'x' }] }).reason;
    // independence from moves and resizes, listings, discard, clear
    const es3 = m.createEditState(); es3.commitText(sc, id('ar-mixed'), { paras: [{ p: 0, runs: [{ r: 3, text: runs[3] + '!' }] }] }); es3.commit(sc, id('ar-mixed'), 10, 5); es3.commitResize(sc, id('ar-wrapped'), { w: el('ar-wrapped').world.w / E + 20, h: el('ar-wrapped').world.h / E, tx: 10, ty: 0 });
    out.mix = { size: es3.size, moves: es3.list().length, resizes: es3.listResizes().length, texts: es3.listTexts().length, has: es3.has(id('ar-mixed')), hasText: es3.hasText(id('ar-mixed')), forSlide: es3.textsForSlide(sc.part).length, other: es3.textsForSlide('ppt/slides/slide9.xml').length, frame: !!es3.frame(id('ar-mixed')) };
    es3.discardText(id('ar-mixed')); out.dt = { has: es3.has(id('ar-mixed')), mv: !!es3.get(id('ar-mixed')), size: es3.size };
    es3.commitText(sc, id('ar-mixed'), { paras: [{ p: 0, runs: [{ r: 3, text: 'z' }] }] }); es3.discard(id('ar-mixed')); out.d = { has: es3.has(id('ar-mixed')), text: es3.hasText(id('ar-mixed')), mv: !!es3.get(id('ar-mixed')), size: es3.size };
    es3.commitText(sc, id('ar-mixed'), { paras: [{ p: 0, runs: [{ r: 3, text: 'z' }] }] }); es3.clear(); out.cl = { size: es3.size, texts: es3.listTexts().length };
    const esA = m.createEditState(); const esB = m.createEditState(); esA.commitText(sc, id('ar-mixed'), { paras: [{ p: 0, runs: [{ r: 3, text: 'q' }] }] }); out.sep = { a: esA.listTexts().length, b: esB.listTexts().length };
    const cp = esA.getText(id('ar-mixed')); cp.paras[0].runs[0].text = 'CHANGED'; out.copy = esA.getText(id('ar-mixed')).paras[0].runs[0].text === 'q';
    return out;
  }, { plain: b64(AR), empty: emptyDeck, br: brDeck });
  ok(us.c1.ok && us.c1.changed && us.c1.keys === 'id,paras,part,slideIndex' && us.c1.plain && us.c1.paras.length === 1 && us.c1.paras[0].runs.length === 1 && us.c1.paras[0].runs[0].r === 3 && us.c1.paras[0].runs[0].text.endsWith('!') && us.c1.paras[0].runs[0].was.length + 1 === us.c1.paras[0].runs[0].text.length && us.c1.v === 1 && us.c1.size === 1, 'commitText keeps a plain-data record { id, part, slideIndex, paras }: only the changed run (r = 3) with its new text and the file\'s text as `was`; a run given with its file text is dropped', us.c1);
  ok(us.same.ok && us.same.changed === false && us.same.v === 1, 'committing the same text again changes nothing (version unchanged)', us.same);
  ok(us.c3.changed && JSON.stringify(us.c3.runs.map((x) => x[0])) === '[0,3]' && us.c3.runs[0][1] === '' && us.c3.runs[0][2].length > 0 && us.c3.n === 1 && us.c3.v === 2, 'a new commit is the TOTAL since the file: it replaces the record, runs sorted by index, a run emptied entirely is kept as an empty run (text "")', us.c3);
  ok(us.eff.texts[0] === '' && us.eff.texts[3].endsWith('!') && us.eff.fmt && us.eff.sceneSame && us.eff.otherSame && us.eff.body && us.eff.newObj, 'effectiveText: only `text` differs (every other field of every run is equal), it is made of new objects, the shape\'s own e.text is not written, an element without a record IS its e.text', us.eff);
  ok(us.stale === true, 'a record whose `was` no longer matches the file text is not applied (the run is returned as it is)');
  ok(us.back.ok && us.back.edit === null && us.back.changed && us.back.n === 0 && us.back.size === 0 && us.back.get === null && us.noneAgain.ok && us.noneAgain.changed === false, 'committing the file\'s own text removes the record: the element has the file\'s text again; removing what is not there changes nothing', { back: us.back, none: us.noneAgain });
  ok(JSON.stringify(us.ref) === JSON.stringify({ table: 'not-editable', noDraft: 'invalid', noParas: 'invalid', badP: 'invalid', badR: 'invalid', type: 'invalid', dup: 'invalid', fillOnRuns: 'invalid', unknown: 'not-editable', noScene: 'not-editable' }) && us.afterRef.v === us.afterRef.v0 && us.afterRef.n === 0, 'refusals, each with its reason, and nothing is written: a table, no / malformed draft, a paragraph or run that is not there, a non-string, a run given twice, a fill for a paragraph that has runs, an unknown id, no scene', { ref: us.ref, after: us.afterRef });
  ok(us.fill.ok && JSON.stringify(us.fill.rec) === JSON.stringify([{ p: 1, runs: [], fill: { text: 'وسط' } }]) && JSON.stringify(us.fill.runs) === JSON.stringify([['وسط', true, 22, true, true]]) && us.fill.p0same && us.fill.p2same && us.fill.sceneSame, 'text for an EMPTY paragraph (fill): kept as { text }; the effective paragraph gets one run with the paragraph\'s default formatting (size 22, right-to-left, src.def); the neighbouring paragraphs are the very same objects; the scene still has the empty paragraph', us.fill);
  ok(us.fillEmpty.ok && us.fillGone === 0 && us.fillBad === 'invalid', 'an empty fill removes it; a fill for a paragraph with text is refused', { e: us.fillEmpty.ok, g: us.fillGone, b: us.fillBad });
  ok(us.mix.size === 2 && us.mix.moves === 1 && us.mix.resizes === 1 && us.mix.texts === 1 && us.mix.has && us.mix.hasText && us.mix.forSlide === 1 && us.mix.other === 0 && us.mix.frame, 'text, move and resize records are independent: one element with a text and a move, another with a resize (size 2), listings and frame() unaffected', us.mix);
  ok(us.dt.has === true && us.dt.mv === true && us.d.has === false && us.d.text === false && us.d.mv === false && us.cl.size === 0 && us.cl.texts === 0, 'discardText() forgets the text only (the move stays); discard() forgets text, move and resize; clear() forgets everything', { dt: us.dt, d: us.d, cl: us.cl });
  ok(us.sep.a === 1 && us.sep.b === 0 && us.copy, 'one edit-state per opened deck keeps its own texts, and records are handed out as copies', { sep: us.sep });

  // ---- 3. real keyboard on p31-arabic: Latin in an Arabic run, bidi marks, NBSP; formatting and structure kept ------------------------------------------
  await newE2(); await drawB(b64(AR), 1); const fmt0 = await fmt('ar-mixed'); const pre = await te('state');
  await beginN('ar-mixed'); await page.evaluate(() => globalThis.__te.caretAt(0, 4)); await kb.insertText('abc'); await outside();
  const c = await page.evaluate(() => { const E = globalThis.__E; const rec = E.listTexts()[0]; const e = [...globalThis.__s.sc.elements].find((x) => x.name === 'ar-mixed'); return { rec, orig: e.text.paras[0].runs.map((r) => r.text), sel: globalThis.__s.sel.id, ev: globalThis.__s.editEvents.slice(-1)[0] }; });
  const post = await te('state'); const dom = await domText('ar-mixed'); const fmt1 = await fmt('ar-mixed');
  ok(c.rec && c.rec.paras.length === 1 && c.rec.paras[0].runs.length === 1 && c.rec.paras[0].runs[0].r === 0 && c.rec.paras[0].runs[0].text === c.orig[0].slice(0, 4) + 'abc' + c.orig[0].slice(4) && c.rec.paras[0].runs[0].was === c.orig[0], 'Latin typed into the first (Arabic) run and a press outside: ONE run changed, its text is the file\'s with "abc" inserted exactly where the caret was, `was` is the file text', c.rec);
  ok(c.ev.phase === 'end' && c.ev.reason === 'outside' && c.ev.committed === true && c.ev.changed === true && c.sel === null, 'onEdit reported the end with reason "outside", committed and changed', c.ev);
  ok(dom.length === 1 && dom[0] === c.orig[0].slice(0, 4) + 'abc' + c.orig[0].slice(4) + c.orig.slice(1).join(''), 'the slide shows the kept text (the whole paragraph is the file\'s text with "abc" inserted)', dom);
  ok(JSON.stringify(fmt1) === JSON.stringify(fmt0), 'and it LOOKS the same apart from the text: every run keeps its lang, font size, family, colour, weight, style and decoration (computed styles of all four runs identical)', { fmt0, fmt1 });
  ok(post.model === pre.model && post.xml === pre.xml, 'the Scene Model (resolved text, frames) and the XML documents are exactly as before: only the edit-state holds the change');
  const kd = await page.evaluate(() => { const n = [...globalThis.__s.hnd.root.querySelectorAll('.sf-el')].find((x) => x.dataset.name === 'ar-mixed'); const ps = [...n.querySelectorAll('.sf-p')]; return { paras: ps.length, dir: ps[0].dir, spans: n.querySelectorAll('.sf-r').length, ce: document.querySelectorAll('[contenteditable]').length }; });
  ok(kd.paras === 1 && kd.dir === 'rtl' && kd.spans === 4 && kd.ce === 0, 'one paragraph, still dir="rtl", the same four runs, nothing editable left', kd);
  // the same drawn again from the edit-state alone (a redraw) is the very same DOM
  const html1 = await page.evaluate(() => [...globalThis.__s.hnd.root.querySelectorAll('.sf-el')].find((x) => x.dataset.name === 'ar-mixed').outerHTML);
  await drawB(b64(AR), 1); const html2 = await page.evaluate(() => [...globalThis.__s.hnd.root.querySelectorAll('.sf-el')].find((x) => x.dataset.name === 'ar-mixed').outerHTML);
  ok(html1 === html2 && html2.includes('abc'), 'a redraw of the slide draws the kept text, and the DOM is byte-identical to the one the commit produced (same builder, one code path for both)');
  await drawB(b64(AR), 0); const o0 = await page.evaluate(() => globalThis.__E.listTexts().length); await drawB(b64(AR), 1); const back = await domText('ar-mixed');
  ok(o0 === 1 && back[0].includes('abc'), 'another slide and back: the kept text is still there (records belong to their slide part)');
  // a second edit builds on the kept text; going back to the file's text removes the record
  await beginN('ar-mixed'); const t2 = (await te('shape')).text[0]; ok(t2.includes('abc'), 'editing the shape again starts from the kept text');
  await page.evaluate(() => globalThis.__te.caretAt(0, 4)); for (let k = 0; k < 3; k++) await kb.press('Delete'); await outside();
  const gone = await page.evaluate(() => ({ n: globalThis.__E.listTexts().length, v: globalThis.__E.version })); const dom3 = await domText('ar-mixed');
  ok(gone.n === 0 && dom3[0] === c.orig.join(''), 'deleting "abc" again (back to the file\'s text) removes the record and the slide shows the file\'s text', { gone, dom3 });

  // bidi marks and NBSP: kept verbatim in the record and on the slide
  await newE2(); await drawB(marks, 1); const preM = await te('state'); await beginN('ar-mixed');
  await page.evaluate(() => globalThis.__te.caretAt(0, 3)); await kb.insertText('\u200F\u00A0\u200E\u202B\u202C'); await outside();
  const rm = await page.evaluate(() => globalThis.__E.listTexts()[0]); const dm = await domText('ar-mixed'); const postM = await te('state');
  ok(rm.paras[0].runs[0].text.includes('\u200F\u00A0\u200E\u202B\u202C') && rm.paras[0].runs[0].was.startsWith('\u200F') && dm[0].includes('tex\u00A02.5\u200E') && dm[0].includes('\u200F\u00A0\u200E\u202B\u202C') && dm[0].startsWith('\u200F'), 'RLM, NBSP, LRM, RLE and PDF typed into the text, and the RLM / NBSP / LRM the file already had, are all in the record and on the slide exactly as characters (nothing normalised, nothing dropped, nothing moved)', { text: rm.paras[0].runs[0].text.split('').map((x) => x.charCodeAt(0).toString(16)).join(' ').slice(0, 80) });
  ok(postM.model === preM.model && postM.xml === preM.xml, 'and the scene and the XML are unchanged');
}

// ---- 4. bullets: the text of a bulleted paragraph is kept, bullets / levels / paragraph count are not touched ---------------------------------------
{
  const AR = path.join(FX, 'p31-arabic.pptx'); const TX = path.join(FX, 'p31-text.pptx'); const kb = page.keyboard;
  const te = (fn, ...a) => page.evaluate(({ fn, a }) => globalThis.__te[fn](...a), { fn, a });
  const newE2 = () => page.evaluate(() => globalThis.__newE());
  const beginN = (name) => page.evaluate((name) => globalThis.__s.sel.beginEdit(globalThis.__te.pick(name)), name);
  const outside = () => page.evaluate(() => { globalThis.__ptr('pointerdown', 4, 4); globalThis.__ptr('pointerup', 4, 4); });
  const shape = (name) => page.evaluate((name) => { const n = [...globalThis.__s.hnd.root.querySelectorAll('.sf-el')].find((x) => x.dataset.name === name); const ps = [...n.querySelectorAll('.sf-text-in > .sf-p')]; return { n: ps.length, bu: ps.map((p) => { const b = p.querySelector(':scope > .sf-bu'); return b ? b.textContent + '|' + b.style.cssText : null; }), dir: ps.map((p) => p.dir), pad: ps.map((p) => p.style.cssText), text: ps.map((p) => [...p.querySelectorAll('.sf-r')].map((r) => r.textContent).join('')) }; }, name);
  for (const [file, idx, name, label] of [[AR, 0, 'Content Placeholder 2', 'Arabic RTL list'], [TX, 0, 'bullets', 'English list with levels']]) {
    await newE2(); await draw(file, idx, true); const before = await shape(name); const pre = await te('state'); await beginN(name);
    await page.evaluate(() => globalThis.__te.caret(1, 'end')); await kb.type(' EDITED'); await page.evaluate(() => globalThis.__te.caretAt(0, 2)); await kb.insertText('Q'); await outside();
    const after = await shape(name); const rec = await page.evaluate(() => globalThis.__E.listTexts()[0]); const post = await te('state');
    ok(rec && rec.paras.length === 2 && rec.paras[0].p === 0 && rec.paras[1].p === 1 && rec.paras.every((p) => p.runs.length === 1), `${label}: two paragraphs edited, two runs changed, nothing else is in the record`, rec && rec.paras);
    ok(after.text[1] === before.text[1] + ' EDITED' && after.text[0] === before.text[0].slice(0, 2) + 'Q' + before.text[0].slice(2) && after.text.slice(2).join('|') === before.text.slice(2).join('|'), `${label}: the slide shows the edited text in those two paragraphs and the others are untouched`, after.text);
    ok(after.n === before.n && JSON.stringify(after.bu) === JSON.stringify(before.bu) && JSON.stringify(after.dir) === JSON.stringify(before.dir) && JSON.stringify(after.pad) === JSON.stringify(before.pad), `${label}: the same paragraphs, every bullet (its character and its inline style), every paragraph direction, indent and spacing exactly as before`, { before: before.bu, after: after.bu });
    ok(post.model === pre.model && post.xml === pre.xml, `${label}: scene and XML unchanged`);
  }
  // auto-numbered bullets keep counting (the paragraph structure did not change)
  await newE2(); await draw(TX, 0, true); const nb = await page.evaluate(() => [...globalThis.__s.hnd.root.querySelectorAll('.sf-el')].map((n) => [n.dataset.name, [...n.querySelectorAll('.sf-bu')].map((b) => b.textContent).join(',')]).filter((x) => x[1]));
  await beginN('bullets'); await page.evaluate(() => globalThis.__te.caret(2, 'end')); await kb.type('!'); await outside(); const nb2 = await page.evaluate(() => [...globalThis.__s.hnd.root.querySelectorAll('.sf-el')].map((n) => [n.dataset.name, [...n.querySelectorAll('.sf-bu')].map((b) => b.textContent).join(',')]).filter((x) => x[1]));
  ok(JSON.stringify(nb) === JSON.stringify(nb2) && nb.length >= 1, 'bullet labels (characters, numbers) of every shape on the slide are identical after a text commit', { nb, nb2 });
}

// ---- 5. empty paragraphs: kept, editable, typed text gets the paragraph's default formatting ------------------------------------------------------------
{
  const AR = path.join(FX, 'p31-arabic.pptx'); const kb = page.keyboard;
  const te = (fn, ...a) => page.evaluate(({ fn, a }) => globalThis.__te[fn](...a), { fn, a });
  const newE2 = () => page.evaluate(() => globalThis.__newE());
  const drawB = (b, idx) => page.evaluate(({ b, idx }) => globalThis.__draw(b, idx, true), { b, idx });
  const beginN = (name) => page.evaluate((name) => globalThis.__s.sel.beginEdit(globalThis.__te.pick(name)), name);
  const outside = () => page.evaluate(() => { globalThis.__ptr('pointerdown', 4, 4); globalThis.__ptr('pointerup', 4, 4); });
  const RPR = '<a:rPr lang="ar-SA" sz="2200"><a:solidFill><a:srgbClr val="FFFFFF"/></a:solidFill><a:latin typeface="Arial"/><a:cs typeface="Arial"/></a:rPr>';
  const emptyDeck = await page.evaluate(({ b, RPR }) => globalThis.__patchSlide(b, 'ppt/slides/slide2.xml', [['<a:t>نص داخل شكل</a:t></a:r></a:p></p:txBody>', '<a:t>نص داخل شكل</a:t></a:r></a:p><a:p><a:pPr algn="ctr" rtl="1"/><a:endParaRPr lang="ar-SA" sz="2200"><a:solidFill><a:srgbClr val="FFFFFF"/></a:solidFill><a:latin typeface="Arial"/><a:cs typeface="Arial"/></a:endParaRPr></a:p><a:p><a:pPr algn="ctr" rtl="1"/><a:r>' + RPR + '<a:t>سطر ثالث</a:t></a:r></a:p></p:txBody>']]), { b: b64(AR), RPR });
  const look = () => page.evaluate(() => { const n = [...globalThis.__s.hnd.root.querySelectorAll('.sf-el')].find((x) => x.dataset.name === 'ar-shape'); return [...n.querySelectorAll('.sf-text-in > .sf-p')].map((p) => { const r = p.querySelector('.sf-r'); const c = getComputedStyle(r); return { dir: p.dir, text: r.textContent.replace('\u200b', ''), fs: c.fontSize, ff: c.fontFamily, color: c.color, lang: r.lang, h: Math.round(p.getBoundingClientRect().height) }; }); });
  await newE2(); await drawB(emptyDeck, 1); const l0 = await look(); const pre = await te('state'); await beginN('ar-shape');
  await page.evaluate(() => globalThis.__te.caret(1, 'start')); await kb.insertText('وسط جديد'); await outside();
  const rec = await page.evaluate(() => globalThis.__E.listTexts()[0]); const l1 = await look(); const post = await te('state');
  ok(l0.length === 3 && l0[1].text === '' && rec && rec.paras.length === 1 && rec.paras[0].p === 1 && rec.paras[0].runs.length === 0 && rec.paras[0].fill.text === 'وسط جديد', 'text typed into the EMPTY paragraph is kept as a fill for paragraph 1 (the three paragraphs are still three)', { l0: l0.length, rec: rec && rec.paras });
  ok(l1.length === 3 && l1[1].text === 'وسط جديد' && l1[1].dir === 'rtl' && l1[1].fs === l1[0].fs && l1[1].ff === l1[0].ff && l1[1].color === l1[0].color && l1[0].text === 'نص داخل شكل' && l1[2].text === 'سطر ثالث', 'it is drawn with the paragraph\'s DEFAULT formatting: the same font size, family and colour as the text paragraph above it (what its <a:endParaRPr> and the inheritance say), the paragraph stays RTL, its neighbours are unchanged', { l0: l0[1], l1: l1[1], ref: l1[0] });
  ok(post.model === pre.model && post.xml === pre.xml, 'the scene still has the empty paragraph (no run) and the XML is unchanged');
  await drawB(emptyDeck, 1); const l2 = await look(); ok(JSON.stringify(l2) === JSON.stringify(l1), 'a redraw draws the same (the fill comes back from the edit-state)');
  // edit it again: delete the text -> the paragraph is empty again, the record is gone
  await beginN('ar-shape'); const sh = (await te('shape')); ok(sh.text.join('|') === 'نص داخل شكل|وسط جديد|سطر ثالث', 'the shape opens for editing with the filled paragraph as text', sh.text);
  await page.evaluate(() => globalThis.__te.caret(1, 'end')); for (let k = 0; k < 8; k++) await kb.press('Backspace'); await outside();
  const l3 = await look(); const n3 = await page.evaluate(() => globalThis.__E.listTexts().length);
  ok(n3 === 0 && l3.length === 3 && l3[1].text === '' && l3[1].h === l0[1].h, 'deleting the text again removes the record: the paragraph is empty again and has the height it had in the file', { n3, l3: l3[1], l0: l0[1] });
  // typing, deleting part of it, typing more in the same session
  await beginN('ar-shape'); await page.evaluate(() => globalThis.__te.caret(1, 'start')); await kb.insertText('abc'); await kb.press('Backspace'); await kb.insertText('XY'); await outside();
  const rec2 = await page.evaluate(() => globalThis.__E.listTexts()[0]); ok(rec2 && rec2.paras[0].fill.text === 'abXY', 'typing, deleting and typing again inside the empty paragraph in one session keeps the final text', rec2 && rec2.paras);
  // a paragraph that is NOT empty and typed over completely keeps its run (and its formatting), it is not a fill
  await beginN('ar-shape'); await page.evaluate(() => { const t = globalThis.__te.inner().children[0].querySelector('.sf-r').firstChild; const r = document.createRange(); r.setStart(t, 0); r.setEnd(t, 3); const g = getSelection(); g.removeAllRanges(); g.addRange(r); }); await kb.insertText('كل'); await outside();
  const rec3 = await page.evaluate(() => globalThis.__E.listTexts()[0]); ok(rec3.paras.some((p) => p.p === 0 && p.runs.length === 1 && p.runs[0].r === 0 && !p.fill) && rec3.paras.some((p) => p.p === 1 && p.fill), 'text typed over a paragraph that has a run stays in that run (formatting of the run), while the empty paragraph keeps its fill', rec3.paras);
}

// ---- 6. line breaks (<a:br>): kept where they are; each line's text goes to its own run ----------------------------------------------------------------
{
  const AR = path.join(FX, 'p31-arabic.pptx'); const kb = page.keyboard;
  const te = (fn, ...a) => page.evaluate(({ fn, a }) => globalThis.__te[fn](...a), { fn, a });
  const newE2 = () => page.evaluate(() => globalThis.__newE());
  const drawB = (b, idx) => page.evaluate(({ b, idx }) => globalThis.__draw(b, idx, true), { b, idx });
  const beginN = (name) => page.evaluate((name) => globalThis.__s.sel.beginEdit(globalThis.__te.pick(name)), name);
  const outside = () => page.evaluate(() => { globalThis.__ptr('pointerdown', 4, 4); globalThis.__ptr('pointerup', 4, 4); });
  const RPR = '<a:rPr lang="ar-SA" sz="2200"><a:solidFill><a:srgbClr val="FFFFFF"/></a:solidFill><a:latin typeface="Arial"/><a:cs typeface="Arial"/></a:rPr>';
  const brDeck = await page.evaluate(({ b, RPR }) => globalThis.__patchSlide(b, 'ppt/slides/slide2.xml', [['<a:t>نص داخل شكل</a:t></a:r>', '<a:t>سطر أول</a:t></a:r><a:br><a:rPr lang="ar-SA" sz="2200"/></a:br><a:r>' + RPR + '<a:t>سطر ثان</a:t></a:r>']]), { b: b64(AR), RPR });
  const dom = () => page.evaluate(() => { const n = [...globalThis.__s.hnd.root.querySelectorAll('.sf-el')].find((x) => x.dataset.name === 'ar-shape'); const p = n.querySelector('.sf-text-in > .sf-p'); return { kids: [...p.childNodes].map((c) => c.tagName === 'BR' ? 'BR' : c.textContent), br: p.querySelectorAll('br').length, paras: n.querySelectorAll('.sf-text-in > .sf-p').length }; });
  const info = await page.evaluate(async (b) => { const { parsed, scene } = await globalThis.__load(b); const sc = await scene.buildScene(parsed, 1); await (await import('/js/model/resolve.js')).resolveScene(sc); const e = sc.elements.find((x) => x.name === 'ar-shape'); return e.text.paras[0].runs.map((r) => r.br ? 'BR' : [r.text, r.src.ci, r.src.s]); }, brDeck);
  ok(JSON.stringify(info) === JSON.stringify([['سطر أول', 1, 0], 'BR', ['سطر ثان', 3, 0]]), 'a paragraph with a line break resolves to run, br, run; the runs\' src.ci are the <a:r> children (1 and 3: the <a:br> is child 2)', info);
  await newE2(); await drawB(brDeck, 1); const pre = await te('state'); const d0 = await dom(); await beginN('ar-shape');
  const t0 = (await te('shape')).text[0];
  // the line break cannot be removed: Backspace at the start of line 2, Delete at the end of line 1, a selection across it
  await page.evaluate(() => { const i = globalThis.__te.inner(); const p = i.children[0]; const spans = p.querySelectorAll('.sf-r'); const r = document.createRange(); r.setStart(spans[1].firstChild, 0); r.collapse(true); const g = getSelection(); g.removeAllRanges(); g.addRange(r); });
  await kb.press('Backspace');
  await page.evaluate(() => { const i = globalThis.__te.inner(); const p = i.children[0]; const spans = p.querySelectorAll('.sf-r'); const r = document.createRange(); r.setStart(spans[0].firstChild, spans[0].firstChild.data.length); r.collapse(true); const g = getSelection(); g.removeAllRanges(); g.addRange(r); });
  await kb.press('Delete');
  await page.evaluate(() => { const i = globalThis.__te.inner(); const p = i.children[0]; const spans = p.querySelectorAll('.sf-r'); const r = document.createRange(); r.setStart(spans[0].firstChild, 2); r.setEnd(spans[1].firstChild, 2); const g = getSelection(); g.removeAllRanges(); g.addRange(r); });
  await kb.press('Backspace'); await kb.insertText('Z'); await kb.press('Control+a'); await kb.press('Backspace');
  const mid = await page.evaluate(() => ({ editing: globalThis.__s.sel.editing, br: globalThis.__te.inner() ? globalThis.__te.inner().querySelectorAll('br').length : -1, text: globalThis.__s.sel.editText() }));
  ok(mid.editing && mid.br === 1 && mid.text[0] === t0, 'Backspace at the start of line 2, Delete at the end of line 1, a selection across the break (deleted or typed over) and select-all + Backspace all leave the line break and the text alone', mid);
  // edit each line: each line's text goes to its own run
  await page.evaluate(() => { const p = globalThis.__te.inner().children[0]; const spans = p.querySelectorAll('.sf-r'); const g = getSelection(); const r1 = document.createRange(); r1.setStart(spans[0].firstChild, 3); r1.collapse(true); g.removeAllRanges(); g.addRange(r1); });
  await kb.insertText('1'); await page.evaluate(() => { const p = globalThis.__te.inner().children[0]; const spans = p.querySelectorAll('.sf-r'); const g = getSelection(); const r2 = document.createRange(); r2.setStart(spans[1].firstChild, spans[1].firstChild.data.length); r2.collapse(true); g.removeAllRanges(); g.addRange(r2); }); await kb.insertText('2'); await outside();
  const rec = await page.evaluate(() => globalThis.__E.listTexts()[0]); const d1 = await dom(); const post = await te('state');
  ok(rec && rec.paras[0].runs.length === 2 && rec.paras[0].runs[0].r === 0 && rec.paras[0].runs[1].r === 2 && rec.paras[0].runs[0].text === 'سطر1 أول' && rec.paras[0].runs[1].text === 'سطر ثان2', 'a character typed in line 1 and one in line 2: two runs changed, r = 0 and r = 2 (the break is entry 1 of the paragraph\'s runs and is never in the record)', rec && rec.paras[0].runs);
  ok(d1.br === 1 && d1.paras === 1 && JSON.stringify(d1.kids) === JSON.stringify(['سطر1 أول', 'BR', 'سطر ثان2']) && d0.br === 1, 'the slide shows both lines edited with the one line break still between them', d1);
  ok(post.model === pre.model && post.xml === pre.xml, 'scene and XML unchanged');
  // a draft that is not intact (a line break added behind the rules' back) is not kept
  await newE2(); await drawB(brDeck, 1); await page.evaluate(() => globalThis.__te.pick('ar-shape')); const pre2 = await te('state'); await beginN('ar-shape');
  await page.evaluate(() => { const i = globalThis.__te.inner(); i.children[0].appendChild(document.createElement('br')); i.dispatchEvent(new Event('input', { bubbles: true })); });
  const gone = await page.evaluate(() => ({ editing: globalThis.__s.sel.editing, n: globalThis.__E.listTexts().length, ev: globalThis.__s.editEvents.slice(-1)[0] })); const post2 = await te('state');
  ok(gone.editing === null && gone.n === 0 && gone.ev.reason === 'structure' && gone.ev.committed === false && post2.html === pre2.html, 'an added line break is a structure violation: the mode ends ("structure"), nothing is kept, the original DOM is back', gone);
}

// ---- 7. runs that are deleted, text that lands outside the original runs, the script / font limitation ---------------------------------------------------
{
  const AR = path.join(FX, 'p31-arabic.pptx'); const kb = page.keyboard;
  const te = (fn, ...a) => page.evaluate(({ fn, a }) => globalThis.__te[fn](...a), { fn, a });
  const newE2 = () => page.evaluate(() => globalThis.__newE());
  const drawB = (b, idx) => page.evaluate(({ b, idx }) => globalThis.__draw(b, idx, true), { b, idx });
  const beginN = (name) => page.evaluate((name) => globalThis.__s.sel.beginEdit(globalThis.__te.pick(name)), name);
  const outside = () => page.evaluate(() => { globalThis.__ptr('pointerdown', 4, 4); globalThis.__ptr('pointerup', 4, 4); });
  const sel2 = (k1, k2) => page.evaluate(({ k1, k2 }) => { const para = globalThis.__te.inner().children[0]; const w = document.createTreeWalker(para, NodeFilter.SHOW_TEXT, { acceptNode: (n) => (n.parentElement.closest('.sf-bu') ? NodeFilter.FILTER_REJECT : NodeFilter.FILTER_ACCEPT) }); const at = (k) => { let n; let left = k; const w2 = document.createTreeWalker(para, NodeFilter.SHOW_TEXT); while ((n = w2.nextNode())) { if (left <= n.data.length) return [n, left]; left -= n.data.length; } return null; }; void w; const a = at(k1); const b = at(k2); const r = document.createRange(); r.setStart(a[0], a[1]); r.setEnd(b[0], b[1]); const g = getSelection(); g.removeAllRanges(); g.addRange(r); }, { k1, k2 });
  const orig = await page.evaluate(async (b) => { const { parsed, scene } = await globalThis.__load(b); const sc = await scene.buildScene(parsed, 1); await (await import('/js/model/resolve.js')).resolveScene(sc); return sc.elements.find((e) => e.name === 'ar-mixed').text.paras[0].runs.map((r) => r.text); }, b64(AR));
  const fmt = () => page.evaluate(() => { const n = [...globalThis.__s.hnd.root.querySelectorAll('.sf-el')].find((x) => x.dataset.name === 'ar-mixed'); return [...n.querySelectorAll('.sf-text-in .sf-r')].map((r) => { const c = getComputedStyle(r); return [r.lang, c.fontSize, c.fontFamily, c.color].join('|'); }); });
  // (a) a whole run deleted: it stays a run, with no text
  await newE2(); await drawB(b64(AR), 1); const f0 = await fmt(); await beginN('ar-mixed'); await sel2(0, orig[0].length); await kb.press('Backspace'); await outside();
  let rec = await page.evaluate(() => globalThis.__E.listTexts()[0]); let eff = await page.evaluate(() => { const e = [...globalThis.__s.sc.elements].find((x) => x.name === 'ar-mixed'); return globalThis.__E.effectiveText(e).paras[0].runs.map((r) => r.text); });
  ok(rec && rec.paras[0].runs.length === 1 && rec.paras[0].runs[0].r === 0 && rec.paras[0].runs[0].text === '' && eff.length === 4 && eff.slice(1).join('') === orig.slice(1).join(''), 'deleting all the text of a run keeps the run (r = 0, text ""): the paragraph still has its four runs, the other three untouched', { runs: rec && rec.paras[0].runs, eff });
  ok(JSON.stringify((await fmt()).slice(1)) === JSON.stringify(f0.slice(1)), 'and the other runs look exactly as before (lang, size, family, colour)');
  // (b) typing right after that: whichever node the browser puts the new text in, the paragraph text is exactly what was typed and no run was lost
  await newE2(); await drawB(b64(AR), 1); await beginN('ar-mixed'); await sel2(0, orig[0].length); await kb.press('Backspace'); await kb.insertText('E'); await outside();
  rec = await page.evaluate(() => globalThis.__E.listTexts()[0]); eff = await page.evaluate(() => { const e = [...globalThis.__s.sc.elements].find((x) => x.name === 'ar-mixed'); return globalThis.__E.effectiveText(e).paras[0].runs.map((r) => r.text); });
  const dom = await page.evaluate(() => [...[...globalThis.__s.hnd.root.querySelectorAll('.sf-el')].find((x) => x.dataset.name === 'ar-mixed').querySelectorAll('.sf-r')].map((r) => r.textContent).join(''));
  ok(eff.length === 4 && eff.join('') === 'E' + orig.slice(1).join('') && dom === eff.join('') && rec.paras[0].runs.every((r) => r.was === orig[r.r]), 'a run deleted and a character typed in its place: the kept paragraph text is exactly "E" + the other runs, still four runs, every `was` is the file\'s text of that run, the slide shows it', { eff, dom });
  // (c) typed text outside every original run (the browser made a span / a text node): attached to the run before it, never lost
  await newE2(); await drawB(b64(AR), 1); await beginN('ar-mixed');
  await page.evaluate(() => { const p = globalThis.__te.inner().children[0]; const spans = [...p.querySelectorAll('.sf-r')]; const stray = document.createElement('span'); stray.append(document.createTextNode('+S')); spans[1].after(stray); const bare = document.createTextNode('+B'); p.append(bare); });
  await outside(); rec = await page.evaluate(() => globalThis.__E.listTexts()[0]);
  ok(rec && rec.paras[0].runs.length === 2 && rec.paras[0].runs[0].r === 1 && rec.paras[0].runs[0].text === orig[1] + '+S' && rec.paras[0].runs[1].r === 3 && rec.paras[0].runs[1].text === orig[3] + '+B', 'text that is not inside an original run goes to the original run before it (a span after run 1, a bare text node at the end): "+S" to run 1, "+B" to the last run, nothing lost', rec && rec.paras[0].runs);
  // (d) the documented limitation: Arabic typed into a Latin run keeps THAT run's formatting (its font stack, no lang)
  await newE2(); await drawB(b64(AR), 1); const f1 = await fmt(); await beginN('ar-mixed'); await page.evaluate(() => globalThis.__te.caretAt(0, 8 + 3)); await kb.insertText('مرحبا'); await outside();
  rec = await page.evaluate(() => globalThis.__E.listTexts()[0]); const f2 = await fmt(); const dm = await page.evaluate(() => [...globalThis.__s.hnd.root.querySelectorAll('.sf-el')].find((x) => x.dataset.name === 'ar-mixed').querySelector('.sf-p').dir);
  ok(rec.paras[0].runs.length === 1 && rec.paras[0].runs[0].r === 1 && rec.paras[0].runs[0].text === orig[1].slice(0, 3) + 'مرحبا' + orig[1].slice(3) && JSON.stringify(f2) === JSON.stringify(f1) && dm === 'rtl', 'Arabic typed into the Latin run "Nabtex 2.5" stays in that run with ITS formatting: same lang (none), same Latin font stack, same size and colour (documented limitation: the script is not re-segmented, so the complex-script font is not applied to it); the paragraph is still RTL and the browser orders the text by the bidi rules', { runs: rec.paras[0].runs, f1, f2 });
}

// ---- 8. cancel keeps what was kept; moves, resizes and text together; reset ------------------------------------------------------------------------------
{
  const AR = path.join(FX, 'p31-arabic.pptx'); const kb = page.keyboard;
  const te = (fn, ...a) => page.evaluate(({ fn, a }) => globalThis.__te[fn](...a), { fn, a });
  const newE2 = () => page.evaluate(() => globalThis.__newE());
  const drawB = (b, idx) => page.evaluate(({ b, idx }) => globalThis.__draw(b, idx, true), { b, idx });
  const beginN = (name) => page.evaluate((name) => globalThis.__s.sel.beginEdit(globalThis.__te.pick(name)), name);
  const outside = () => page.evaluate(() => { globalThis.__ptr('pointerdown', 4, 4); globalThis.__ptr('pointerup', 4, 4); });
  const node = (name) => page.evaluate((name) => { const n = [...globalThis.__s.hnd.root.querySelectorAll('.sf-el')].find((x) => x.dataset.name === name); return { html: n.querySelector('.sf-text').outerHTML, tr: n.style.translate, tf: n.style.transform, text: n.textContent }; }, name);
  const E = () => page.evaluate(() => ({ t: globalThis.__E.listTexts().length, m: globalThis.__E.list().length, r: globalThis.__E.listResizes().length, size: globalThis.__E.size, v: globalThis.__E.version }));
  await newE2(); await drawB(b64(AR), 1); await page.evaluate(() => globalThis.__te.pick('ar-mixed')); const orig = await node('ar-mixed');
  await beginN('ar-mixed'); await page.evaluate(() => globalThis.__te.caretAt(0, 4)); await kb.insertText('abc'); await outside();
  const kept = await node('ar-mixed'); const e1 = await E(); const v1 = e1.v;
  // Escape after a commit: only the new typing is thrown away
  await beginN('ar-mixed'); await kb.insertText('NOPE'); await kb.press('Escape'); const afterEsc = await node('ar-mixed'); const e2 = await E();
  ok(kept.text.includes('abc') && afterEsc.html === kept.html && e2.v === v1 && e2.t === 1, 'Escape after a commit throws away the new typing only: the kept text, its DOM and the edit-state (version) are exactly as they were', { v1, v2: e2.v });
  // a structure violation likewise
  await beginN('ar-mixed'); await kb.insertText('NOPE'); await page.evaluate(() => { const i = globalThis.__te.inner(); i.appendChild(document.createElement('div')); i.dispatchEvent(new Event('input', { bubbles: true })); }); const afterStruct = await node('ar-mixed'); const e3 = await E();
  ok(afterStruct.html === kept.html && e3.v === v1, 'and so does a structure violation');
  // move after the text commit, redraw
  await page.evaluate(() => globalThis.__moveBy('ar-mixed', 20, 10)); const e4 = await E();
  await drawB(b64(AR), 1); const both = await node('ar-mixed'); const e5 = await E();
  ok(e4.m === 1 && e4.t === 1 && e4.size === 1 && both.html === kept.html && both.tr !== '', 'a move of the text-edited shape is a separate record (1 text, 1 move, 1 element); after a redraw the shape is drawn with both: the kept text and the moved place', { e4, tr: both.tr });
  // editing a moved shape works; the text commit keeps the move
  await beginN('ar-mixed'); await page.evaluate(() => globalThis.__te.caretAt(0, 0)); await kb.insertText('M'); await outside(); const moved = await node('ar-mixed'); const e6 = await E();
  ok(moved.text.includes('Mabc') || moved.text.includes('M'), 'a moved shape can still be edited') ; ok(e6.m === 1 && e6.t === 1 && moved.tr === both.tr, 'and the edit keeps the move (same place, one move record, one text record)', { e6, tr: moved.tr });
  // a resize: allowed, the kept text is drawn stretched; editing is refused while it is resized
  const rz = await page.evaluate(() => { const r = globalThis.__resizeBy('ar-mixed', 'se', 20, 10); const id = globalThis.__s.sel.id; return { dw: r && r.dw, res: globalThis.__s.sel.beginEdit(id) }; }); const withRz = await node('ar-mixed'); const e7 = await E();
  ok(rz.dw === 20 && rz.res.ok === false && rz.res.reason === 'resized' && e7.r === 1 && e7.t === 1 && e7.m === 1 && withRz.text === moved.text && /scale\(/.test(withRz.tf), 'resizing the shape that has a kept text is allowed (text, move and resize are three records of one element: the text is drawn stretched); editing it is refused ("resized") and the text is unchanged', { rz, e7 });
  await drawB(b64(AR), 1); const re = await node('ar-mixed'); const e8 = await E();
  ok(re.text === moved.text && re.tr === moved.tr && /scale\(/.test(re.tf) && e8.size === 1 && e8.t === 1 && e8.m === 1 && e8.r === 1, 'after a redraw all three come back: the kept text, the move, the stretch');
  // reset() undoes all three
  const rs = await page.evaluate(() => { globalThis.__te.pick('ar-mixed'); globalThis.__s.sel.reset(); return null; }); void rs; const e9 = await E(); const fin = await node('ar-mixed');
  ok(e9.size === 0 && e9.t === 0 && fin.html === orig.html && fin.tr === '' && !/scale\(/.test(fin.tf), 'reset() forgets the text, the move and the resize of the slide: the text box DOM is byte-identical to the file\'s and the shape is at its file place and size', { e9 });
  // reset() also ends an edit in progress without keeping it
  await beginN('ar-mixed'); await kb.insertText('NOPE'); await page.evaluate(() => globalThis.__s.sel.reset()); const e10 = await E(); const fin2 = await node('ar-mixed'); ok(e10.size === 0 && fin2.html === orig.html && !fin2.text.includes('NOPE'), 'reset() during an edit ends it without keeping anything');
}

// ---- 9. every fixture and real deck: edit + keep, nothing but the edit-state changes, reset() undoes it byte for byte --------------------------------------
{
  const decks = [GEO, SEL, path.join(FX, 'p31-shapes.pptx'), path.join(FX, 'p31-text.pptx'), path.join(FX, 'p31-arabic.pptx'), path.join(FX, 'p31-media.pptx'), path.join(FX, 'p31-unsupported.pptx'), path.join(FX, 'p31-template-decor.pptx'), path.join(FX, 'p22-geometry.pptx'), ...realFiles];
  let editedTotal = 0; let decksWith = 0;
  for (const f of decks) {
    const r = await page.evaluate(async ({ b }) => {
      const { parsed, scene } = await globalThis.__load(b); const rn = await import('/js/render/render.js'); const sl = await import('/js/render/selection.js'); const em = await import('/js/model/edit-state.js'); const tm = await import('/js/render/text-edit.js');
      const ser = (d) => new XMLSerializer().serializeToString(d); const edits = em.createEditState(); const orig = await globalThis.__zipKeyOf(b);
      const out = { slides: 0, edited: 0, kept: 0, shown: 0, undone: 0, mutated: 0, modelChanged: 0, errors: [], zipSame: (await globalThis.__zipKey(parsed.zip)) === orig };
      for (let i = 0; i < parsed.slides.length; i++) {
        try {
          const sc = await scene.buildScene(parsed, i); const docs = [...new Set(sc.all.map((e) => e.el.ownerDocument))]; const before = docs.map(ser);
          const hnd = await rn.renderSlide(sc, document.getElementById('sf-host'), { width: 960 }); const model0 = JSON.stringify(sc.all.map((e) => [e.world, e.xfrm, e.text]));
          let last = null; const sel = sl.attachSelection(hnd, sc, { edits, onEdit: (m) => { if (m.phase === 'end') last = m; } });
          for (const e of sc.elements) {
            if (!sel.select(e.id)) continue; const node = hnd.nodes.get(e.id); if (!tm.textEditability(e, { node, resized: false }).ok) continue; const html0 = node.outerHTML;
            sel.beginEdit(e.id); document.execCommand('insertText', false, 'تجربة x'); const typed = (sel.editText() || []).some((t) => t.includes('تجربة x')); sel.endEdit(); out.edited++;
            const rec = edits.getText(e.id); if (rec ? last.committed : !last.committed) out.kept++; if (typed && rec && node.textContent.includes('تجربة x')) out.shown++;
            sel.reset(); if (!edits.hasText(e.id) && node.outerHTML === html0) out.undone++;
            break;
          }
          sel.dispose(); docs.forEach((d, j) => { if (ser(d) !== before[j]) out.mutated++; }); if (JSON.stringify(sc.all.map((e) => [e.world, e.xfrm, e.text])) !== model0) out.modelChanged++; out.slides++;
        } catch (e) { out.errors.push(i + ': ' + String(e).slice(0, 200)); }
      }
      out.zipAfter = (await globalThis.__zipKey(parsed.zip)) === orig; return out;
    }, { b: b64(f) });
    const nm = path.basename(f); editedTotal += r.edited; if (r.edited) decksWith++;
    ok(r.errors.length === 0 && r.slides > 0 && r.shown === r.edited && r.kept === r.edited && r.undone === r.edited && r.mutated === 0 && r.modelChanged === 0 && r.zipSame && r.zipAfter, `${nm}: ${r.edited} text shape(s) edited and kept: each shows the typed text and is in the edit-state, reset() brings back the exact DOM; the XML, the resolved text, the frames and every part of the file are unchanged`, r);
  }
  ok(editedTotal >= 25 && decksWith >= 5, `the sweep really kept text (${editedTotal} shapes in ${decksWith} decks)`, { editedTotal, decksWith });
}

// ---- 10. code guards ------------------------------------------------------------------------------------------------------------------------------------
{
  const strip = (t) => t.replace(/\/\*[\s\S]*?\*\//g, '').replace(/\/\/.*$/gm, '');
  const rd = (f) => strip(readFileSync(path.join(ROOT, f), 'utf8'));
  const te = rd('js/render/text-edit.js'); const sl = rd('js/render/selection.js'); const es = rd('js/model/edit-state.js'); const rn = rd('js/render/render.js'); const rv = rd('js/model/resolve.js');
  ok(!/\.(text|world|xfrm|paras|runs)\s*=[^=]/.test(te) && !/\.(text|world|xfrm|paras|runs)\s*=[^=]/.test(sl) && !/\.(world|xfrm)\s*=[^=]/.test(es), 'text-edit.js, selection.js and edit-state.js never assign to a scene element\'s text, paragraphs, runs or frames: the Scene Model is only read');
  ok((sl.match(/edits\.commitText\(/g) || []).length === 1 && (sl.match(/handle\.setText\?\./g) || []).length >= 3 && !/\.defRun\b|\.src\.(ci|s|def)\b/.test(sl + rn) && !/\.src\.(ci|s|def)\b/.test(te) && (te.match(/\.defRun\b/g) || []).length === 1, 'selection.js commits text in one place (afterEdit); the renderer and the controller do not read src or defRun; the editor only asks whether an empty paragraph has a default run (once) and never reads src (only the resolver, the edit-state and the font loader use them)');
  ok(/export function effectiveText/.test(es) && !/innerHTML|document\.|querySelector|contenteditable/.test(es) && /src: \{ ci, s: si\+\+ \}/.test(rv.replace(/\s+/g, ' ')) || /src = \{ ci, s: si\+\+ \}/.test(rv.replace(/\s+/g, ' ')), 'the edit-state is plain data and a pure effectiveText; the resolver sets src on every run and defRun on empty paragraphs, additively');
  ok(!/writeXml|serializeToString|XMLSerializer|zip\.file|generateAsync|saveAs|download/.test(te + sl + es), 'nothing writes XML, a zip or a file: no serializer, no zip write, no download in the editor, the controller or the edit-state');
}


// ======================================================================================================
section('Phase 3.2.5-C: the four fixes of the 3.2.5-B audit (emptied paragraph, line-break-only paragraph, autofit of an edited shape, stale fit notes)');
{
  const PC = path.join(FX, 'p325c-fixes.pptx'); const kb = page.keyboard; const cdp = await ctx.newCDPSession(page);
  const te = (fn, ...a) => page.evaluate(({ fn, a }) => globalThis.__te[fn](...a), { fn, a });
  const fresh = async () => { await page.evaluate(() => globalThis.__newE()); await page.evaluate((b) => globalThis.__draw(b, 0, true), b64(PC)); };
  const begin = (name) => page.evaluate((name) => { const id = globalThis.__te.pick(name); return { id, r: globalThis.__s.sel.beginEdit(id) }; }, name);
  const endApi = () => page.evaluate(() => globalThis.__s.sel.endEdit('api'));
  const nEv = () => page.evaluate(() => globalThis.__s.editEvents.length);
  const ends = (n) => page.evaluate((n) => globalThis.__s.editEvents.slice(n).filter((e) => e.phase === 'end').map((e) => ({ reason: e.reason, committed: e.committed, changed: e.changed, skipped: e.skipped })), n);
  const rec = (name) => page.evaluate((name) => { const e = globalThis.__s.sc.elements.find((x) => x.name === name); const r = globalThis.__E.getText(e.id); return r ? r.paras : null; }, name);
  const dom = (name) => page.evaluate((name) => { const n = [...globalThis.__s.hnd.root.querySelectorAll('.sf-el')].find((x) => x.dataset.name === name); return [...n.querySelectorAll('.sf-text-in > .sf-p')].map((p) => ({ dir: p.dir, ce: p.getAttribute('contenteditable'), bu: p.querySelectorAll(':scope > .sf-bu').length, br: p.querySelectorAll(':scope > br').length, spans: [...p.querySelectorAll(':scope > .sf-r')].map((r) => ({ t: r.textContent, ph: r.getAttribute('data-ph'), lang: r.lang, fs: getComputedStyle(r).fontSize, ff: getComputedStyle(r).fontFamily })) })); }, name);
  const box = (name) => page.evaluate((name) => { const n = [...globalThis.__s.hnd.root.querySelectorAll('.sf-el')].find((x) => x.dataset.name === name); const b = n.querySelector('.sf-text'); const i = b.firstElementChild; const r = i.querySelector('.sf-r'); return { autofit: b.dataset.autofit || null, base: b.dataset.fitBase || null, fit: b.style.getPropertyValue('--fit') || null, over: i.scrollHeight - (b.clientHeight - parseFloat(getComputedStyle(b).paddingTop) - parseFloat(getComputedStyle(b).paddingBottom)), fs: parseFloat(getComputedStyle(r).fontSize), support: n.dataset.support || null }; }, name);
  const sup = (name) => page.evaluate((name) => { const e = globalThis.__s.sc.elements.find((x) => x.name === name); return { support: e.support, fit: e.reasons.filter((r) => r.text.startsWith('text overflows')).length, reasons: e.reasons.length, approx: globalThis.__s.hnd.report().counts.approximated }; }, name);
  const emptyPara = async (p) => { await te('caret', p, 'start'); await kb.press('Shift+End'); await kb.press('Backspace'); };
  const html = () => page.evaluate(() => { globalThis.__s.sel.clear(); return globalThis.__s.hnd.root.outerHTML.replace(/<div class="sf-sel"[\s\S]*$/, ''); }); // (the selection's own marks are not the drawing)
  const diffAt = (a, b) => { let i = 0; while (i < a.length && a[i] === b[i]) i++; return { at: i, a: a.slice(Math.max(0, i - 60), i + 120), b: b.slice(Math.max(0, i - 60), i + 120) }; };
  const LONG = ' and then a good deal more text that certainly cannot fit the frame at the original size so it must be fitted';

  let d, r, n, en;
  try {
  // ---- fix 1: the renderer's placeholder span is not a run ----------------------------------------------------------------------------------------
  await fresh();
  const d0 = await dom('single');
  ok(d0[0].spans.length === 1 && d0[0].spans[0].ph === null, 'a paragraph with text carries no placeholder span', d0);
  await begin('single'); await emptyPara(0); await endApi();
  d = await dom('single'); r = await rec('single');
  ok(r && r[0].runs.length === 1 && r[0].runs[0].text === '' && d[0].spans.length === 2 && d[0].spans[0].ph === null && d[0].spans[1].ph === '1' && d[0].spans[1].t === '\u200b', 'fix 1: emptying the only run keeps the record (text "") and the redraw is [the run\'s empty span, the renderer\'s placeholder span marked data-ph]', { r, d });
  n = await nEv(); await begin('single'); await te('caret', 0, 'end'); await kb.insertText('NEW'); await endApi();
  en = await ends(n); r = await rec('single'); d = await dom('single');
  ok(en.length === 1 && en[0].committed === true && en[0].changed === true && r && r[0].runs[0].text === 'NEW' && r[0].runs[0].was === 'Delete me all' && d[0].spans.length === 1 && d[0].spans[0].t === 'NEW' && d[0].spans[0].fs === d0[0].spans[0].fs && d[0].spans[0].ff === d0[0].spans[0].ff, 'fix 1: re-entering the emptied paragraph and typing KEEPS the text (it used to be dropped without a message), in the same run: same size and font', { en, r, d });
  ok(!JSON.stringify(await rec('single')).includes('\u200b'), 'the placeholder\'s zero-width character never becomes text in the record');

  await fresh(); await begin('single'); await emptyPara(0); await endApi();
  n = await nEv(); await begin('single'); const cEmpty = await te('shape'); await endApi(); en = await ends(n); r = await rec('single');
  ok(cEmpty.text[0] === '' && en[0].changed === false && r[0].runs[0].text === '' && !JSON.stringify(r).includes('\u200b'), 'entering the emptied paragraph and leaving it untouched changes nothing (still ""), no zero-width character leaks into the record', { cEmpty, en, r });

  await fresh(); await begin('twopara'); await emptyPara(1); await endApi();
  n = await nEv(); await begin('twopara'); await te('caret', 0, 'end'); await kb.insertText(' EDIT'); await endApi(); en = await ends(n); r = await rec('twopara');
  ok(en[0].committed && en[0].changed && r[0].runs[0].text === 'First paragraph EDIT' && r[1].runs[0].text === '', 'fix 1: after one paragraph was emptied, an edit of ANOTHER paragraph of the same shape is kept (the whole shape used to refuse every later edit)', { en, r });
  n = await nEv(); await begin('twopara'); await te('caret', 1, 'end'); await kb.insertText('Y'); await endApi(); en = await ends(n); r = await rec('twopara');
  ok(en[0].committed && en[0].changed && r[1].runs[0].text === 'Y' && r[0].runs[0].text === 'First paragraph EDIT', 'fix 1: and typing into the emptied paragraph itself is kept too, the other paragraph\'s edit stays', { en, r });

  await fresh(); await begin('bul'); await emptyPara(0); await endApi(); d = await dom('bul');
  ok(d.length === 2 && d.every((p) => p.bu === 1) && d[0].spans[0].t === '', 'an emptied bulleted paragraph keeps its bullet', d);
  n = await nEv(); await begin('bul'); await te('caret', 0, 'end'); await kb.insertText('Z'); await endApi(); en = await ends(n); d = await dom('bul'); r = await rec('bul');
  ok(en[0].committed && en[0].changed && r[0].runs[0].text === 'Z' && d.every((p) => p.bu === 1) && d[0].spans.map((x) => x.t).join('') === 'Z' && d[1].spans[0].t === 'Second bullet', 'fix 1 + bullets: typing into the emptied bulleted paragraph is kept, both bullets and the other paragraph are intact', { en, r, d });

  await fresh(); await begin('rtl'); await emptyPara(0); await endApi();
  n = await nEv(); await begin('rtl'); await te('caret', 0, 'end'); await kb.insertText('نص'); await endApi(); en = await ends(n); d = await dom('rtl'); r = await rec('rtl');
  ok(en[0].committed && en[0].changed && r[0].runs[0].text === 'نص' && d[0].dir === 'rtl' && d[0].spans.length === 1 && d[0].spans[0].lang === 'ar' && d[0].spans[0].fs === '28px', 'fix 1 + RTL: typing into an emptied right-to-left paragraph is kept; direction and language stay', { en, r, d });

  await fresh(); await begin('single'); await emptyPara(0); await endApi();
  n = await nEv(); await begin('single'); await te('caret', 0, 'end'); await cdp.send('Input.imeSetComposition', { text: 'あ', selectionStart: 1, selectionEnd: 1 }); await cdp.send('Input.imeSetComposition', { text: 'あい', selectionStart: 2, selectionEnd: 2 }); await cdp.send('Input.insertText', { text: '愛' });
  const stillIme = await page.evaluate(() => globalThis.__s.sel.editing); await endApi(); en = await ends(n); r = await rec('single');
  ok(stillIme && en[0].committed && en[0].changed && r[0].runs[0].text === '愛', 'fix 1 + IME: a composition into the emptied paragraph is kept as well (the mode goes on during it, the text is the committed one)', { stillIme, en, r });

  } catch (e) { ok(false, 'a check of the group above threw: ' + String(e).slice(0, 300)); }
  try {
  // ---- fix 2: paragraphs of line breaks only are not editable ----------------------------------------------------------------------------------------------
  await fresh(); const html0 = await html(); const dB = await dom('bronly');
  ok(dB.length === 4 && dB[1].br === 1 && dB[1].spans.length === 0 && dB[2].br === 2 && dB[0].ce === null, 'the fixture: paragraphs 1 and 2 hold line breaks only (1 and 2 breaks), no run', dB);
  n = await nEv(); const b1 = await begin('bronly'); const dIn = await dom('bronly'); const startEv = await page.evaluate((n) => globalThis.__s.editEvents[n], n);
  ok(b1.r.ok === true && dIn.map((p) => p.ce).join() === ',false,false,' && startEv.locked === 2, 'fix 2: entering the mode makes exactly the two line-break-only paragraphs non-editable (contenteditable="false"), the start event says 2 were locked', { dIn, startEv });
  const textBefore = await page.evaluate(() => globalThis.__s.sel.editText()); await te('caret', 0, 'end'); await kb.insertText(' EARLIER');
  await page.evaluate(() => { const i = globalThis.__te.inner(); const p = i.children[1]; const rg = document.createRange(); rg.setStart(p, 0); rg.collapse(true); const g = getSelection(); g.removeAllRanges(); g.addRange(rg); });
  await kb.insertText('x'); await kb.press('Backspace'); await kb.press('Delete');
  await cdp.send('Input.imeSetComposition', { text: 'か', selectionStart: 1, selectionEnd: 1 }); await cdp.send('Input.insertText', { text: '漢' });
  const stillOn = await page.evaluate(() => globalThis.__s.sel.editing); const textNow = await page.evaluate(() => globalThis.__s.sel.editText());
  ok(stillOn && textNow[0] === 'Line one EARLIER' && JSON.stringify(textNow.slice(1)) === JSON.stringify(textBefore.slice(1)), 'fix 2: keys and an IME composition aimed at a line-break-only paragraph do nothing, the session goes on (it used to end with "structure" and throw the whole draft away)', { stillOn, textNow });
  await te('caret', 3, 'end'); await kb.insertText('!'); await endApi(); en = await ends(n); r = await rec('bronly'); const dAfter = await dom('bronly');
  ok(en.length === 1 && en[0].reason === 'api' && en[0].committed === true && en[0].changed === true && r && r.length === 2 && r[0].runs[0].text === 'Line one EARLIER' && r[1].runs[0].text === 'Line four!', 'fix 2: the edits made before and after the attempt are KEPT, only those two paragraphs are in the record', { en, r });
  ok(dAfter.every((p) => p.ce === null) && dAfter[1].br === 1 && dAfter[2].br === 2 && dAfter[1].spans.length === 0 && (await page.evaluate(() => document.querySelectorAll('[contenteditable]').length)) === 0, 'after the mode the break-only paragraphs are untouched (same breaks, no run) and no contenteditable is left anywhere', dAfter);
  await page.evaluate(() => globalThis.__s.sel.reset()); { const hh = await html(); ok(hh === html0, 'reset() brings back the exact DOM of the file, break-only paragraphs included', diffAt(hh, html0)); }

  await fresh(); n = await nEv(); await begin('brtext'); const bt = await dom('brtext'); await te('caretAt', 0, 6); await kb.insertText('!'); const sBt = await page.evaluate(() => globalThis.__s.editEvents[globalThis.__s.editEvents.length - 1]); await endApi(); en = await ends(n); r = await rec('brtext');
  ok(bt[0].ce === null && bt[0].br === 1 && bt[0].spans.length === 2 && en[0].committed && en[0].changed && r[0].runs.some((x) => x.text === 'Line A!'), 'a paragraph with text AND a line break is still editable (only break-ONLY paragraphs are locked)', { bt, en, r });
  const ref = await begin('allbr');
  ok(ref.r.ok === false && ref.r.reason === 'line-breaks-only' && (await page.evaluate(() => globalThis.__s.sel.editing)) === null, 'a shape whose only paragraph is a line break is refused (line-breaks-only): nothing to type into', ref);

  await fresh(); n = await nEv(); await begin('brlast'); await kb.insertText('!'); const bl = await page.evaluate(() => globalThis.__s.sel.editText()); await endApi(); en = await ends(n); r = await rec('brlast');
  ok(bl[0] === 'Last text!' && en[0].committed && en[0].changed && r.length === 1 && r[0].runs[0].text === 'Last text!', 'fix 2: when the last paragraph holds only a line break, the caret of a fresh edit goes into the last paragraph that CAN be typed into', { bl, en, r });
  await fresh(); n = await nEv(); await begin('bronly');
  await page.evaluate(() => { const i = globalThis.__te.inner(); const p = i.children[1]; const rg = document.createRange(); rg.setStart(p, 0); rg.collapse(true); const g = getSelection(); g.removeAllRanges(); g.addRange(rg); const dt = new DataTransfer(); dt.setData('text/plain', 'pasted'); i.dispatchEvent(new ClipboardEvent('paste', { clipboardData: dt, bubbles: true, cancelable: true })); });
  const pasted = await page.evaluate(() => [globalThis.__s.sel.editing, globalThis.__s.sel.editText()]); await endApi(); en = await ends(n);
  ok(pasted[0] && !pasted[1].join('').includes('pasted') && en[0].changed === false, 'a paste aimed at a line-break-only paragraph inserts nothing and leaves the session alone', pasted);

  } catch (e) { ok(false, 'a check of the group above threw: ' + String(e).slice(0, 300)); }
  try {
  // ---- fix 3: an edited shape with a stored fontScale is fitted again --------------------------------------------------------------------------------------
  await fresh(); const html1 = await html(); const s0 = await box('scaled');
  ok(s0.autofit === null && s0.fit === null && s0.base === null && Math.abs(s0.fs - 20) < 0.05 && s0.over <= 1, 'unedited: the shape is drawn at the file\'s stored scale (20 px = 32 pt x 62.5 %), no autofit, no overflow (unchanged)', s0);
  await begin('scaled'); await te('caret', 0, 'end'); await kb.insertText(LONG + LONG + LONG); await endApi();
  const s1 = await box('scaled');
  ok(s1.autofit === 'shrink' && s1.base === '0.625' && s1.fit !== null && s1.over <= 1 && s1.fs < 20, 'fix 3: after adding a lot of text the shape shrinks again below the stored scale and no longer overflows (it overflowed by ~190 px)', s1);
  await begin('scaled'); await te('caret', 0, 'start'); await kb.press('Shift+End'); await page.evaluate(() => { const i = globalThis.__te.inner(); const rg = document.createRange(); rg.selectNodeContents(i.children[0]); const g = getSelection(); g.removeAllRanges(); g.addRange(rg); }); await kb.insertText('Short.'); await endApi();
  const s2 = await box('scaled');
  ok(s2.over <= 1 && Math.abs(s2.fs - 32) < 0.1 && s2.support !== 'approximated', 'fix 3: shortened text grows back to the full size (32 px) - it stayed at the stored 62.5 % before - and carries no fit note', s2);
  await begin('scaled'); await page.evaluate(() => { const i = globalThis.__te.inner(); const rg = document.createRange(); rg.selectNodeContents(i.children[0]); const g = getSelection(); g.removeAllRanges(); g.addRange(rg); }); await kb.insertText(LONG + LONG); await endApi();
  const h2 = await html(); await page.evaluate((b) => globalThis.__draw(b, 0, true), b64(PC)); const h3 = await html();
  ok(h2 === h3, 'a redraw from the edit-state reproduces the committed DOM byte for byte (autofit state included)', diffAt(h2, h3));
  await page.evaluate(() => globalThis.__s.sel.reset()); { const hh = await html(); ok(hh === html1, 'reset() brings back the file\'s DOM byte for byte: stored scale again, no data-autofit, no --fit', diffAt(hh, html1)); }
  await begin('scaled'); await te('caret', 0, 'end'); await kb.insertText('X'); await endApi(); ok((await rec('scaled')) !== null, 'one more character is a record');
  await begin('scaled'); await te('caret', 0, 'end'); await kb.press('Backspace'); await endApi(); { const hh = await html(); ok((await rec('scaled')) === null && hh === html1, 'typing the file\'s text back removes the record and the shape is drawn as the file has it (stored scale)', diffAt(hh, html1)); }

  await fresh(); const k0 = await box('shrink');
  await begin('shrink'); await te('caret', 0, 'end'); await kb.insertText(LONG); await endApi(); const k1 = await box('shrink');
  ok(k0.autofit === 'shrink' && k0.base === null && k0.fit === null && k1.fit !== null && k1.over <= 1 && k1.base === null, 'a shape with normAutofit and no stored scale still shrinks when text is added (behaviour unchanged, no fit base)', { k0, k1 });
  await begin('shrink'); await te('caret', 0, 'end'); for (let i = 0; i < LONG.length; i++) await kb.press('Backspace'); await endApi(); const k2 = await box('shrink');
  ok(k2.fit === null && Math.abs(k2.fs - k0.fs) < 0.01, 'and returns to the full size when the text is back to the file\'s', { k2, k0 });

  } catch (e) { ok(false, 'a check of the group above threw: ' + String(e).slice(0, 300)); }
  try {
  // ---- notes written by anything else survive the clearing of the fit note
  await fresh(); await page.evaluate(async () => { const m = await import('/js/model/scene.js'); const e = globalThis.__s.sc.elements.find((x) => x.name === 'shrink'); m.note(e, 'approximated', 'another reason'); });
  await begin('shrink'); await te('caret', 0, 'end'); await kb.insertText(LONG); await endApi(); const o1 = await sup('shrink');
  await begin('shrink'); await te('caret', 0, 'end'); for (let i = 0; i < LONG.length; i++) await kb.press('Backspace'); await endApi(); const o2 = await sup('shrink');
  ok(o1.fit === 1 && o1.reasons === 2 && o2.fit === 0 && o2.reasons === 1 && o2.support === 'approximated', 'a note of another origin stays (and keeps the support level) when the fit note is taken off', { o1, o2 });

  } catch (e) { ok(false, 'a check of the group above threw: ' + String(e).slice(0, 300)); }
  try {
  // ---- fix 4: no stale fit note, the details panel follows the commit ---------------------------------------------------------------------------------------
  await fresh(); const u0 = await sup('shrink');
  ok(u0.fit === 0 && u0.support === 'native', 'the fixture\'s shrink shape starts with no fit note', u0);
  await begin('shrink'); await te('caret', 0, 'end'); await kb.insertText(LONG); await endApi(); const u1 = await sup('shrink'); const nd1 = await box('shrink');
  await begin('shrink'); await te('caret', 0, 'end'); await kb.insertText(LONG); await endApi(); const u2 = await sup('shrink');
  ok(u1.fit === 1 && u1.support === 'approximated' && u1.approx === u0.approx + 1 && nd1.support === 'approximated' && u2.fit === 1 && u2.reasons === u1.reasons, 'fix 4: a fit note is written once; a second edit REPLACES it (a note per percentage piled up before)', { u0, u1, u2 });
  await page.evaluate(() => globalThis.__s.sel.reset()); const u3 = await sup('shrink'); const nd3 = await box('shrink');
  ok(u3.fit === 0 && u3.support === 'native' && u3.approx === u0.approx && nd3.support === 'native', 'fix 4: after reset() the note is gone and the element is "native" again (the count in the report is the file\'s)', { u3, nd3 });
  await begin('shrink'); await te('caret', 0, 'end'); await kb.insertText(LONG); await endApi();
  await begin('shrink'); await te('caret', 0, 'end'); for (let i = 0; i < LONG.length; i++) await kb.press('Backspace'); await endApi(); const u4 = await sup('shrink');
  ok((await rec('shrink')) === null && u4.fit === 0 && u4.support === 'native' && u4.approx === u0.approx, 'fix 4: restoring the file\'s text through editing also clears the note and the support level (it stayed "approximated" before)', u4);
  } catch (e) { ok(false, 'a check of the group above threw: ' + String(e).slice(0, 300)); }
}

section('Rendered tab details after a text commit (Phase 3.2.5-C, fix 4), in the real UI');
{
  const ui = await ctx.newPage(); watch(ui, 'details-ui: ');
  await ui.setViewportSize({ width: 1280, height: 900 });
  await openRendered(ui, 'Fix deck', path.join(FX, 'p325c-fixes.pptx'));
  const info = () => ui.evaluate(() => { const t = [...document.querySelectorAll('.rendered p.muted.fine')].map((x) => x.textContent).find((x) => /drawn from the file/.test(x)) || ''; const m = /(\d+) approximated/.exec(t); return { approx: m ? Number(m[1]) : null, line: t, details: document.querySelector('.rendered .details')?.textContent || '' }; });
  const pt = (name) => ui.evaluate((n) => { const el = [...document.querySelectorAll('.rendered-host .sf-el')].find((x) => x.dataset.name === n); const r = (el.querySelector('.sf-text-in') || el).getBoundingClientRect(); return { x: r.left + r.width / 2, y: r.top + r.height / 2 }; }, name);
  const a0 = await info(); const q = await pt('shrink');
  ok(a0.approx !== null && !/shrink on overflow/.test(a0.details), 'the details start without a fit note', a0);
  await ui.mouse.click(q.x, q.y); await ui.mouse.dblclick(q.x, q.y); await ui.keyboard.press('End');
  await ui.keyboard.type(' and then a good deal more text that certainly cannot fit the frame at the original size so it must be fitted');
  await ui.mouse.click(...(await at(ui, 900, 520)));
  await ui.waitForFunction(() => !document.querySelector('.rendered-host [contenteditable]'));
  const a1 = await info();
  ok(a1.approx === a0.approx + 1 && /shrink on overflow/.test(a1.details), 'fix 4: after the commit the count line and the details list the overflow note (they were never refreshed before)', { a0: a0.approx, a1: a1.approx, line: a1.line, d: a1.details.slice(0, 200) });
  await ui.mouse.dblclick(q.x, q.y); await ui.keyboard.press('Control+A'); await ui.keyboard.type('Autofit shrink box with a modest amount of text here');
  await ui.mouse.click(...(await at(ui, 900, 520))); await ui.waitForFunction(() => !document.querySelector('.rendered-host [contenteditable]'));
  const a2 = await info();
  ok(a2.approx === a0.approx && !/shrink on overflow/.test(a2.details), 'fix 4: typing the file\'s text back brings the count and the details back to the file\'s state (no stale note)', { a0: a0.approx, a2: a2.approx, line: a2.line });
  await ui.close();
}


// ======================================================================================================
section('Phase 3.2.6-A: style edit-state (text colour / size / font / bold / italic, fill and line colour) - edit-state.js only');
{
  const P326 = path.join(FX, 'p326a-styles.pptx');
  await page.evaluate((b) => globalThis.__draw(b, 0), b64(P326)); // only the scene is used: the edit-states below are created by the tests
  const res = await page.evaluate(async () => {
    const es = await import('/js/model/edit-state.js');
    const sc = globalThis.__s.sc; const out = []; const J = (v) => JSON.stringify(v);
    const T = (name, cond, detail) => out.push([name, !!cond, cond ? undefined : (detail === undefined ? '' : J(detail)).slice(0, 600)]);
    const el = (n) => sc.elements.find((x) => x.name === n); const id = (n) => el(n).id;
    const snap = () => J(sc.elements.map((e) => [e.id, e.text, e.style, e.support, e.reasons, e.xfrm]));
    const before = snap();
    const C = (st, n, patch) => st.commitStyle(sc, id(n), patch);
    const same = (a, b) => J(a) === J(b);
    const fresh = () => es.createEditState();
    const fake = (n, over) => { const e = { ...el(n), ...over }; return { byId: new Map([[e.id, e]]), index: 0 }; };

    // ---- the fixture is what the tests assume ----
    const fx = (n) => { const e = el(n); return { fill: e.style && e.style.fill, line: e.style && e.style.line, runs: (e.text?.paras || []).flatMap((p) => p.runs), def: (e.text?.paras || []).map((p) => p.defRun) }; };
    T('fixture: solid = fill FF0000, solid outline 0000FF, run 20 pt 112233 Arial', fx('solid').fill.hex === 'FF0000' && fx('solid').line.type === 'solid' && fx('solid').line.hex === '0000FF' && fx('solid').runs[0].sz === 20 && fx('solid').runs[0].hex === '112233' && fx('solid').runs[0].font === 'Arial', fx('solid'));
    T('fixture: nofill has no fill and no outline, grad a gradient, mixed 4 runs (one a link), scaled fontScale 0.625 with a 20 px run, emptyp no run and a default run of 28 pt', fx('nofill').fill.type === 'none' && fx('nofill').line.type === 'none' && fx('grad').fill.type === 'gradient' && fx('mixed').runs.length === 4 && fx('mixed').runs[3].link && el('scaled').text.body.fontScale === 0.625 && fx('scaled').runs[0].sz === 20 && fx('emptyp').runs.length === 0 && fx('emptyp').def[0].sz === 28, 0);

    // ---- surface and an empty state ----
    const S0 = fresh();
    T('the API exists: commitStyle getStyle hasStyle stylesForSlide listStyles discardStyle effectiveStyle effectiveStyledText', ['commitStyle', 'getStyle', 'hasStyle', 'stylesForSlide', 'listStyles', 'discardStyle', 'effectiveStyle', 'effectiveStyledText'].every((k) => typeof S0[k] === 'function') && typeof es.effectiveStyle === 'function' && typeof es.effectiveText === 'function');
    T('a new state holds no style: size 0, version 0, no record, empty lists, has() false', S0.size === 0 && S0.version === 0 && S0.getStyle(id('solid')) === null && !S0.hasStyle(id('solid')) && !S0.has(id('solid')) && S0.listStyles().length === 0 && S0.stylesForSlide(el('solid').part).length === 0);

    // ---- fill ----
    { const S = fresh(); const r = C(S, 'solid', { fill: { hex: '00ff00' } });
      T('fill: ok, a copy of the record { fill: { hex: "00FF00" } } with what the file had, changed, version 1', r.ok && r.changed && J(Object.keys(r.edit)) === J(['id', 'part', 'slideIndex', 'fill', 'was']) && r.edit.fill.hex === '00FF00' && r.edit.id === id('solid') && r.edit.part === el('solid').part && r.edit.was.fill.type === 'solid' && r.edit.was.fill.hex === 'FF0000' && S.version === 1 && S.size === 1 && S.hasStyle(id('solid')) && S.has(id('solid')), r);
      const r2 = C(S, 'solid', { fill: { hex: '#00FF00' } });
      T('the same fill again (with a leading #, upper case) changes nothing: changed false, version still 1', r2.ok && !r2.changed && S.version === 1 && S.getStyle(id('solid')).fill.hex === '00FF00', r2);
      const r3 = C(S, 'solid', { fill: { hex: 'FF0000' } });
      T('a fill equal to the file\'s removes the record: ok, edit null, changed, version 2, nothing left', r3.ok && r3.edit === null && r3.changed && S.version === 2 && !S.hasStyle(id('solid')) && S.size === 0, r3);
      C(S, 'solid', { fill: { hex: '123456' } }); const v = S.version; const r4 = C(S, 'solid', { fill: null });
      T('fill: null reverts it (record gone, version +1); reverting again is a no-op (changed false, version unchanged)', r4.ok && r4.changed && r4.edit === null && S.version === v + 1 && C(S, 'solid', { fill: null }).changed === false && S.version === v + 1);
      const g = C(S, 'grad', { fill: { hex: '0000FF' } }); const nf = C(S, 'nofill', { fill: { hex: '0000FF' } });
      T('a solid fill replaces a gradient fill and a missing fill (both kept, with what the file had)', g.ok && g.edit.was.fill.type === 'gradient' && g.edit.was.fill.hex === null && nf.ok && nf.edit.was.fill.type === 'none', [g, nf]);
      const ch = C(S, 'grpchild', { fill: { hex: '0000FF' } });
      T('the child of a group can be styled', ch.ok && ch.changed, ch); }

    // ---- line ----
    { const S = fresh(); const r = C(S, 'solid', { line: { hex: 'ffff00' } });
      T('line colour: ok, record { line: { hex: "FFFF00" } }, was { hex 0000FF, alpha 1 }', r.ok && r.changed && r.edit.line.hex === 'FFFF00' && r.edit.was.line.hex === '0000FF' && r.edit.was.line.alpha === 1 && !('fill' in r.edit) && !('text' in r.edit), r);
      T('a line colour equal to the file\'s is not kept; null reverts', C(S, 'solid', { line: { hex: '0000ff' } }).edit === null && !S.hasStyle(id('solid')) && C(S, 'solid', { line: { hex: '000000' } }).changed && C(S, 'solid', { line: null }).edit === null && !S.hasStyle(id('solid')));
      const c = C(S, 'conn', { line: { hex: '123456' } });
      T('a connector line\'s colour can be edited', c.ok && c.edit.was.line.hex === 'AA00AA', c);
      const f = C(S, 'conn', { fill: { hex: '123456' } }); const t = C(S, 'conn', { text: { b: true } });
      T('a connector has no fill and no text: no-fill-target, no-text', !f.ok && f.reason === 'no-fill-target' && !t.ok && t.reason === 'no-text', [f, t]);
      const n = C(S, 'nofill', { line: { hex: '123456' } }); const m = C(S, 'mixed', { line: { hex: '123456' } });
      T('no outline (type none) or no shape outline at all: no-outline', !n.ok && n.reason === 'no-outline' && !m.ok && m.reason === 'no-outline', [n, m]); }

    // ---- text: set, merge, revert ----
    { const S = fresh(); const r = C(S, 'solid', { text: { color: 'abcdef', size: 36, font: 'Georgia', b: true, i: true } });
      T('text: all five properties are kept, in a fixed order, with what the file had (color 112233, size 20, font Arial, b false, i false)', r.ok && J(r.edit.text) === J({ color: 'ABCDEF', size: 36, font: 'Georgia', b: true, i: true }) && J(r.edit.was.text) === J({ color: '112233', size: 20, font: 'Arial', b: false, i: false }), r);
      const r2 = C(S, 'solid', { text: { size: 40.255 } });
      T('a patch merges into the record (the other properties stay) and the size has two decimals', r2.ok && r2.edit.text.size === 40.26 && r2.edit.text.color === 'ABCDEF' && r2.edit.text.b === true, r2);
      const r3 = C(S, 'solid', { text: { color: null, b: null } });
      T('null reverts a single property only', r3.ok && J(Object.keys(r3.edit.text)) === J(['size', 'font', 'i']) && !('color' in r3.edit.was.text) && !('b' in r3.edit.was.text), r3);
      const r4 = C(S, 'solid', { text: { size: 20, font: 'arial', i: false } });
      T('values equal to the file (size 20, font "arial" in another case, italic false) drop out of the record; with nothing else left there is none', r4.ok && r4.changed && r4.edit === null && !S.hasStyle(id('solid')), r4);
      T('text: null reverts every text property at once', C(S, 'solid', { text: { size: 22, b: true } }).ok && S.hasStyle(id('solid')) && C(S, 'solid', { text: null }).edit === null && !S.hasStyle(id('solid')));
      const p = C(S, 'solid', { text: { b: false } });
      T('bold false equals the file (not bold): nothing is kept', p.ok && p.edit === null && !S.hasStyle(id('solid')), p); }

    // ---- mixed runs, links, stored fontScale, empty paragraph ----
    { const S = fresh(); const m = C(S, 'mixed', { text: { size: 18, b: true, i: false, color: 'FF0000' } });
      T('mixed runs: a value only SOME runs have is kept (size 18, bold, red), one every run has is dropped (italic false is not true of the italic run: kept too)', m.ok && m.edit.text.size === 18 && m.edit.text.b === true && m.edit.text.i === false && m.edit.text.color === 'FF0000' && m.edit.was.text.size === null && m.edit.was.text.b === null && m.edit.was.text.color === null, m);
      const t = S.effectiveStyledText(el('mixed')).paras[0].runs;
      T('effective text: every run has the style; the hyperlink run keeps the theme\'s colour (and gets size / bold / italic)', t.every((r) => r.sz === 18 && r.b === true && r.i === false) && t.slice(0, 3).every((r) => r.hex === 'FF0000' && r.alpha === 1) && t[3].hex === '0000FF' && t[3].link === true, t);
      const S2 = fresh(); const u = C(S2, 'mixed', { text: { b: false } });
      T('bold false is not what every run has (one is bold): kept', u.ok && u.edit && u.edit.text.b === false, u);
      const lk = fake('mixed', { text: { ...el('mixed').text, paras: [{ ...el('mixed').text.paras[0], runs: [el('mixed').text.paras[0].runs[3]] }] } });
      const nc = fresh().commitStyle(lk, id('mixed'), { text: { color: 'FF0000' } }); const ns = fresh().commitStyle(lk, id('mixed'), { text: { size: 30 } });
      T('a text colour for text that is only hyperlinks: no-color-target; its other properties are fine', !nc.ok && nc.reason === 'no-color-target' && ns.ok && ns.changed, [nc, ns]); }
    { const t0 = el('scaled').text; const rd = fake('scaled', { text: { ...t0, body: { ...t0.body, fontScale: 0.925 }, paras: [{ ...t0.paras[0], runs: [{ ...t0.paras[0].runs[0], sz: 16.7 }] }] } }); // 18 pt x 0.925 = 16.65, rounded to 16.7 by the resolver
      const S = fresh(); const a = S.commitStyle(rd, id('scaled'), { text: { size: 18 } }); const b = S.commitStyle(rd, id('scaled'), { text: { size: 18.5 } });
      T('the resolver\'s rounding of a scaled size (16.7 for 18 x 0.925) does not make 18 look like a change; 18.5 is one', a.ok && a.edit === null && b.ok && b.edit && b.edit.text.size === 18.5, [a, b]); }
    { const S = fresh(); const a = C(S, 'scaled', { text: { size: 32 } }); const b = C(S, 'scaled', { text: { size: 30 } });
      T('stored fontScale: the nominal size is compared (run 20 / 0.625 = 32 equals the file: no record), 30 is kept', a.ok && a.edit === null && b.ok && b.edit.text.size === 30 && b.edit.was.text.size === 32, [a, b]);
      const t = S.effectiveStyledText(el('scaled')).paras[0].runs[0];
      T('...and the effective run size is nominal x fontScale (18.75), exactly what the resolver would have given a run of 30 pt', Math.abs(t.sz - 18.75) < 1e-9, t.sz); }
    { const S = fresh(); const a = C(S, 'emptyp', { text: { size: 28 } }); const b = C(S, 'emptyp', { text: { size: 30, b: true, color: 'FF8800' } });
      T('empty paragraph: its default run (28 pt) is the file value; a style is kept for it', a.ok && a.edit === null && b.ok && b.edit.was.text.size === 28, [a, b]);
      const t = S.effectiveStyledText(el('emptyp')).paras[0];
      T('effective text: the default run has the style', t.defRun.sz === 30 && t.defRun.b === true && t.defRun.hex === 'FF8800' && t.runs.length === 0, t);
      const tx = fresh(); const tr = tx.commitText(sc, id('emptyp'), { paras: [{ p: 0, runs: [], fill: { text: 'typed' } }] }); const sr = tx.commitStyle(sc, id('emptyp'), { text: { size: 33 } });
      const ef = tx.effectiveStyledText(el('emptyp')).paras[0].runs[0];
      T('a text typed into the empty paragraph (a fill run) is drawn with the style as well', tr.ok && sr.ok && ef.text === 'typed' && ef.fill === true && ef.sz === 33 && tx.effectiveText(el('emptyp')).paras[0].runs[0].sz === 28, ef); }

    // ---- unsupported elements and refusals leave everything alone ----
    { const S = fresh(); const v0 = S.version;
      const R = (n, p) => C(S, n, p);
      T('picture, table, group are not styleable', ['pic', 'tbl', 'grp'].every((n) => { const r = R(n, { fill: { hex: '123456' } }); return !r.ok && r.reason === 'not-styleable'; }));
      T('an unknown id and a scene without the element: not-styleable', S.commitStyle(sc, 'nope', { fill: { hex: '123456' } }).reason === 'not-styleable' && S.commitStyle(null, id('solid'), { fill: { hex: '123456' } }).reason === 'not-styleable' && S.commitStyle({}, id('solid'), { fill: { hex: '123456' } }).reason === 'not-styleable');
      T('an element of a layout or master (origin) and an alternate-content element: not-styleable', S.commitStyle(fake('solid', { origin: 'layout' }), id('solid'), { fill: { hex: '123456' } }).reason === 'not-styleable' && S.commitStyle(fake('solid', { kind: 'alt' }), id('solid'), { fill: { hex: '123456' } }).reason === 'not-styleable');
      T('not-styleable comes before invalid', S.commitStyle(sc, id('pic'), 'garbage').reason === 'not-styleable');
      const pf = S.commitStyle(fake('solid', { style: { ...el('solid').style, fill: { type: 'image', target: 'x.png' } } }), id('solid'), { fill: { hex: '123456' } });
      T('a shape whose fill is a picture: picture-fill', !pf.ok && pf.reason === 'picture-fill', pf);
      const nt = C(S, 'grad', { text: { b: true } }); const sn = S.commitStyle(fake('solid', { text: null }), id('solid'), { text: { b: true } });
      T('text properties for a shape with no text: no-text (a gradient shape has text, so it is fine; a text-less copy of a shape is refused)', nt.ok && sn.reason === 'no-text', [nt, sn]);
      T('a refusal changes nothing: version', S.version === v0 + 1); } // the one ok commit above (grad text b)

    // ---- validation: every bad input is 'invalid' and changes nothing ----
    { const S = fresh(); C(S, 'solid', { fill: { hex: '010203' } }); const v0 = S.version; const rec0 = J(S.listStyles());
      const bad = [undefined, null, 'x', 5, [], {}, { text: {} }, { text: 'x' }, { text: [] }, { fill: {} }, { fill: 'FF0000' }, { fill: { hex: 'GGGGGG' } }, { fill: { hex: '12345' } }, { fill: { hex: '1234567' } }, { fill: { hex: '' } }, { fill: { hex: 123456 } }, { fill: { hex: '##123456' } }, { fill: { hex: 'FF0000', alpha: 0.5 } }, { line: { hex: 'xyz' } }, { line: 'red' },
        { text: { color: 'red' } }, { text: { color: 123456 } }, { text: { size: 0 } }, { text: { size: 0.99 } }, { text: { size: -3 } }, { text: { size: 4001 } }, { text: { size: NaN } }, { text: { size: Infinity } }, { text: { size: '12' } }, { text: { size: 1e9 } },
        { text: { font: '' } }, { text: { font: '   ' } }, { text: { font: 5 } }, { text: { font: 'a"b' } }, { text: { font: "a'b" } }, { text: { font: 'a<b' } }, { text: { font: '+mj-lt' } }, { text: { font: 'x'.repeat(101) } }, { text: { font: 'a\u0000b' } },
        { text: { b: 'true' } }, { text: { b: 1 } }, { text: { i: 0 } }, { text: { weight: 700 } }, { text: { weight: true } }, { text: { bold: true } }, { text: { size: 12, extra: 1 } }, { shadow: { hex: 'FF0000' } }, { fill: { hex: 'FF0000' }, shadow: null }, { text: { size: 20, color: 'nope' } }, { fill: { hex: '00FF00' }, line: { hex: 'bad' } }];
      const results = bad.map((p) => { const r = S.commitStyle(sc, id('solid'), p); return r.ok === false && r.reason === 'invalid'; });
      T(`${bad.length} malformed patches and values are all refused as invalid`, results.every(Boolean), bad.filter((_, i) => !results[i]));
      T('...and nothing changed: version, records', S.version === v0 && J(S.listStyles()) === rec0);
      const atomic = S.commitStyle(sc, id('solid'), { fill: { hex: '00FF00' }, text: { size: -1 } });
      T('a patch with one good and one bad part applies neither (atomic)', !atomic.ok && S.version === v0 && J(S.listStyles()) === rec0, atomic);
      const lenient = S.commitStyle(sc, id('solid'), { fill: undefined, text: undefined });
      T('keys with the value undefined count as absent: nothing to do, ok, changed false', lenient.ok && !lenient.changed && S.version === v0, lenient);
      const edge = [C(S, 'solid', { text: { size: 1 } }), C(S, 'solid', { text: { size: 4000 } }), C(S, 'solid', { text: { font: 'A' } }), C(S, 'solid', { text: { font: 'x'.repeat(100) } }), C(S, 'solid', { text: { font: ' Gill Sans MT ' } })];
      T('the limits are inclusive (size 1 and 4000, a 1 and a 100 character font) and a font is trimmed', edge.every((r) => r.ok) && S.getStyle(id('solid')).text.font === 'Gill Sans MT', edge); }

    // ---- combined with moves, resizes and text records ----
    { const S = fresh(); const sid = id('solid');
      const mv = S.commit(sc, sid, 10, 20); const rz = S.commitResize(sc, sid, { w: el('solid').world.w / 12700 + 20, h: el('solid').world.h / 12700, tx: 10, ty: 0 }); const tx = S.commitText(sc, sid, { paras: [{ p: 0, runs: [{ r: 0, text: 'Plain!' }] }] });
      T('move, resize and text commit on the same element', mv.ok && rz.ok && tx.ok, [mv, rz, tx]);
      const a = { move: J(S.get(sid)), resize: J(S.getResize(sid)), text: J(S.getText(sid)), frame: J(S.frame(sid)), list: J(S.list()), rl: J(S.listResizes()), tl: J(S.listTexts()), slide: J(S.forSlide(el('solid').part)), eff: J(S.effectiveText(el('solid'))), v: S.version };
      const st = C(S, 'solid', { fill: { hex: '00FF00' }, line: { hex: 'FFFF00' }, text: { color: 'FFFFFF', size: 30, font: 'Georgia', b: true, i: true } });
      const b = { move: J(S.get(sid)), resize: J(S.getResize(sid)), text: J(S.getText(sid)), frame: J(S.frame(sid)), list: J(S.list()), rl: J(S.listResizes()), tl: J(S.listTexts()), slide: J(S.forSlide(el('solid').part)), eff: J(S.effectiveText(el('solid'))) };
      T('a style commit leaves the move, resize, text records, the frame, every list and effectiveText(e) exactly as they were; version +1', st.ok && st.changed && a.move === b.move && a.resize === b.resize && a.text === b.text && a.frame === b.frame && a.list === b.list && a.rl === b.rl && a.tl === b.tl && a.slide === b.slide && a.eff === b.eff && S.version === a.v + 1, [a, b]);
      const eff = S.effectiveStyledText(el('solid')).paras[0].runs[0];
      T('effectiveStyledText: the edited text AND the style (text "Plain!", 30 pt, FFFFFF, Georgia, bold, italic); effectiveText(e) has the text only', eff.text === 'Plain!' && eff.sz === 30 && eff.hex === 'FFFFFF' && eff.font === 'Georgia' && eff.b && eff.i && S.effectiveText(el('solid')).paras[0].runs[0].sz === 20 && S.effectiveText(el('solid')).paras[0].runs[0].text === 'Plain!', eff);
      const mv2 = S.commit(sc, sid, 30, 40); const tx2 = S.commitText(sc, sid, { paras: [{ p: 0, runs: [{ r: 0, text: 'Again' }] }] });
      T('later moves and text edits leave the style record as it is', mv2.ok && tx2.ok && S.getStyle(sid).fill.hex === '00FF00' && S.getStyle(sid).text.size === 30 && S.getStyle(sid).line.hex === 'FFFF00');
      const v1 = S.version; S.discardStyle(sid);
      T('discardStyle removes the style only (version +1): move, resize, text stay; has() still true; a second discard is a no-op', !S.hasStyle(sid) && S.get(sid) && S.getResize(sid) && S.getText(sid) && S.has(sid) && S.version === v1 + 1 && (S.discardStyle(sid), S.version === v1 + 1));
      C(S, 'solid', { fill: { hex: '00FF00' } }); C(S, 'grad', { fill: { hex: '00FF00' } }); const v2 = S.version;
      S.discardText(sid); T('discardText leaves the styles', S.hasStyle(sid) && S.hasStyle(id('grad')) && S.version === v2 + 1);
      const v3 = S.version; S.discard(sid);
      T('discard(id) forgets move, resize, text AND style of that element, and only of it (version +1)', !S.has(sid) && !S.hasStyle(sid) && S.get(sid) === null && S.getResize(sid) === null && S.getText(sid) === null && S.hasStyle(id('grad')) && S.version === v3 + 1 && S.size === 1);
      S.commit(sc, sid, 5, 5); C(S, 'solid', { fill: { hex: '0000FF' } }); const v4 = S.version; S.clear();
      T('clear() empties moves, resizes, texts and styles (version +1); clearing again does not count', S.size === 0 && S.listStyles().length === 0 && S.list().length === 0 && S.version === v4 + 1 && (S.clear(), S.version === v4 + 1)); }
    { const S = fresh();
      C(S, 'solid', { fill: { hex: '00FF00' } }); S.commit(sc, id('solid'), 5, 5); C(S, 'grad', { text: { b: true } }); S.commitText(sc, id('nofill'), { paras: [{ p: 0, runs: [{ r: 0, text: 'X' }] }] });
      T('size counts every element once, whatever it holds (solid: move + style, grad: style, nofill: text)', S.size === 3, S.size);
      T('stylesForSlide / listStyles: copies, in the order the elements were first styled, only that slide\'s', S.listStyles().map((r) => r.id).join() === [id('solid'), id('grad')].join() && S.stylesForSlide(el('solid').part).length === 2 && S.stylesForSlide('ppt/slides/slide2.xml').length === 0);
      const cp = S.getStyle(id('solid')); cp.fill.hex = '000000'; cp.was.fill.hex = '000000'; const all = S.listStyles(); all[0].fill.hex = '111111';
      T('the records returned are copies: changing them changes nothing in the state', S.getStyle(id('solid')).fill.hex === '00FF00' && S.getStyle(id('solid')).was.fill.hex === 'FF0000');
      const k1 = C(S, 'solid', { fill: { hex: '00ff00' } }).changed; const k2 = C(S, 'grad', { text: { b: true } }).changed;
      T('re-committing the same style (any case) is not a change', k1 === false && k2 === false); }

    // ---- effective style ----
    { const S = fresh(); const e = el('solid');
      T('effectiveStyle without a record IS e.style (same object); the module function too', S.effectiveStyle(e) === e.style && es.effectiveStyle(e, null) === e.style && es.effectiveStyle(null) === null && S.effectiveStyle(null) === null);
      C(S, 'solid', { fill: { hex: '00FF00' }, line: { hex: 'FFFF00' } }); const f = S.effectiveStyle(e);
      T('effectiveStyle with a record: new objects, fill solid 00FF00 alpha 1, line FFFF00 alpha 1 with width / dash / join of the file; e.style untouched', f !== e.style && f.fill !== e.style.fill && f.line !== e.style.line && f.fill.type === 'solid' && f.fill.hex === '00FF00' && f.fill.alpha === 1 && f.line.hex === 'FFFF00' && f.line.alpha === 1 && f.line.w === e.style.line.w && f.line.join === e.style.line.join && e.style.fill.hex === 'FF0000' && e.style.line.hex === '0000FF', f);
      C(S, 'solid', { fill: null }); const g = S.effectiveStyle(e);
      T('only the line edited: the fill object is the file\'s own', g.fill === e.style.fill && g.line.hex === 'FFFF00');
      const sh = { ...el('solid'), style: { ...el('solid').style, shadow: { dx: 1, dy: 2, blur: 3, hex: '000000', alpha: 0.5 }, glow: { rad: 5, hex: 'FFFFFF', alpha: 1 } } };
      const h = es.effectiveStyle(sh, { fill: { hex: '00FF00' } });
      T('shadow and glow are carried over as copies', h.shadow !== sh.style.shadow && J(h.shadow) === J(sh.style.shadow) && h.glow !== sh.style.glow && J(h.glow) === J(sh.style.glow));
      const gr = fresh(); C(gr, 'grad', { fill: { hex: '0000FF' } }); const gf = gr.effectiveStyle(el('grad'));
      T('a gradient replaced by a solid colour has no stops left in the effective fill', gf.fill.type === 'solid' && !('stops' in gf.fill) && el('grad').style.fill.type === 'gradient'); }
    T('module effectiveText(e, rec) with two arguments is what it was: no record gives e.text itself, a record gives new runs', es.effectiveText(el('solid'), null) === el('solid').text && es.effectiveText(el('solid'), undefined, undefined) === el('solid').text && es.effectiveText(el('solid'), null, { fill: { hex: '000000' } }) === el('solid').text);

    T('NOTHING of the Scene Model was written (text, style, support, reasons, xfrm of every element are byte-identical to before)', snap() === before);
    return out;
  });
  for (const [name, pass, detail] of res) ok(pass, name, detail);

  // ---- static guards: edit-state.js only ----
  const src = (f) => readFileSync(path.join(ROOT, f), 'utf8');
  const es = src('js/model/edit-state.js');
  ok(!/\bdocument\b|\bwindow\b|\.style\.|setAttribute|innerHTML|import\(|fetch\(/.test(es.replace(/\/\/.*$/gm, '').replace(/\/\*[\s\S]*?\*\//g, '')), 'edit-state.js stays pure data (no DOM, no CSSOM, no network) after 3.2.6-A');
  ok(['js/render/render.js', 'js/render/selection.js', 'js/render/text-edit.js', 'js/ui/rendered.js', 'js/model/resolve.js', 'js/model/scene.js', 'js/pptx/design.js'].every((f) => !/commitStyle|effectiveStyle|effectiveStyledText|stylesForSlide|discardStyle|hasStyle/.test(src(f))), 'rendering, controller, text editor, UI and resolver do not use the style API yet (that is 3.2.6-B)');
  ok(/commitStyle|effectiveStyledText/.test(es) && !/exportPptx|writeXml|serialize|XMLSerializer|zip\.file/.test(es), 'no XML is written and nothing is exported by the style edit-state');
}

// ======================================================================================================
section('page health');
ok(consoleErrors.length === 0, 'no console errors / CSP violations during the run', consoleErrors.slice(0, 5));

await browser.close(); server.close();
console.log(`\n${pass} passed, ${fail} failed`);
process.exit(fail ? 1 : 0);
