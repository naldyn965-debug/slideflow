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
  globalThis.__draw = async (b64, idx) => {
    const { parsed, scene } = await globalThis.__load(b64);
    const rn = await import('/js/render/render.js'); const sl = await import('/js/render/selection.js');
    const sc = await scene.buildScene(parsed, idx);
    globalThis.__s?.sel.dispose();
    const hnd = await rn.renderSlide(sc, document.getElementById('sf-host'), { width: 960 });
    const changes = []; const moves = [];
    const sel = sl.attachSelection(hnd, sc, { onChange: (id) => changes.push(id), onMove: (m) => moves.push(m.phase) });
    globalThis.__s = { hnd, sc, sel, parsed, changes, moves };
    return { slideWidth: sc.width, ids: sc.elements.map((e) => e.id) };
  };
});
const draw = (file, idx) => page.evaluate(({ b, idx }) => globalThis.__draw(b, idx), { b: b64(file), idx });
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
  // switching slides ends the preview with the slide; coming back shows the slide as in the file
  await ui.click('.rendered .film-item >> nth=1'); await ui.waitForSelector('.rendered-host .sf-slide[data-slide="1"]');
  await ui.click('.rendered .film-item >> nth=0'); await ui.waitForSelector('.rendered-host .sf-slide[data-slide="0"]');
  const back = await mv();
  ok(back.moved === 0 && back.count === 0, 'switching away and back draws the slide as in the file again: the temporary move is gone', back);
  // the same after leaving the Rendered tab and returning
  await ui.mouse.click(...(await at(ui, 450, 300)));
  await drag([450, 300], [500, 330]);
  ok((await mv()).moved === 1, 'moved again');
  await ui.click('role=tab[name="Slides"]'); await ui.click('role=tab[name="Rendered"]'); await ui.waitForSelector('.rendered-host .sf-slide .sf-el');
  const re = await mv();
  ok(re.moved === 0 && re.name === 'front-rect', 'leaving the Rendered tab and coming back: the element is selected again at its file position (the move was temporary)', re);
  // text: dragging a selected text box moves it and does not leave text selected
  await ui.mouse.click(...(await at(ui, 90, 320)));
  ok((await mv()).name === 'textbox', 'the text box is selected');
  await drag([90, 320], [150, 360]);
  const tx = await mv();
  ok(tx.name === 'textbox' && tx.moved === 1 && tx.sel === '', 'dragging a selected text box moves it and selects no text', tx);
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
section('page health');
ok(consoleErrors.length === 0, 'no console errors / CSP violations during the run', consoleErrors.slice(0, 5));

await browser.close(); server.close();
console.log(`\n${pass} passed, ${fail} failed`);
process.exit(fail ? 1 : 0);
