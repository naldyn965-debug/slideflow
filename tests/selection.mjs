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
    const changes = [];
    const sel = sl.attachSelection(hnd, sc, { onChange: (id) => changes.push(id) });
    globalThis.__s = { hnd, sc, sel, parsed, changes };
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
section('page health');
ok(consoleErrors.length === 0, 'no console errors / CSP violations during the run', consoleErrors.slice(0, 5));

await browser.close(); server.close();
console.log(`\n${pass} passed, ${fail} failed`);
process.exit(fail ? 1 : 0);
