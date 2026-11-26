// Phase 3.1 tests: scene model, renderer, fonts, Rendered tab, fidelity and round trip.  Run: node tests/render.mjs
// Runs in headless Chromium (playwright, a dev dependency) against the app served with the real vercel.json headers, exactly like tests/e2e.mjs.
// Sections are added step by step; each prints "ok" / "FAIL" lines and a final count.
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
const ok = (c, msg, extra) => { if (c) { pass++; console.log('  ok   ' + msg); } else { fail++; console.log('FAIL   ' + msg + (extra !== undefined ? '  -> ' + JSON.stringify(extra).slice(0, 400) : '')); } };
const section = (t) => console.log('\n# ' + t);

const FX = path.join(ROOT, 'tests/fixtures');
const REAL = path.join(ROOT, 'tests/real');
const b64 = (f) => readFileSync(f).toString('base64');
const realFiles = [...(existsSync(REAL) ? readdirSync(REAL).filter((f) => f.endsWith('.pptx')).map((f) => path.join(REAL, f)) : []), ...(process.env.REAL_PPTX || '').split(':').filter(Boolean)];

const browser = await chromium.launch();
const ctx = await browser.newContext({ viewport: { width: 1400, height: 900 } });
const page = await ctx.newPage();
const consoleErrors = [];
page.on('console', (m) => {
  if (m.type() !== 'error') return;
  // Same known artefact tests/e2e.mjs ignores: Chrome logs it when parsing slide XML that contains <p:style> elements.
  if (/Applying inline style/.test(m.text()) && m.location().url.endsWith('/js/pptx/parser.js')) return;
  consoleErrors.push(m.text());
});
page.on('pageerror', (e) => consoleErrors.push('pageerror: ' + e.message));
await page.goto(BASE + '/index.html');
await page.waitForFunction(() => !!globalThis.JSZip);

/** Page-side loader (the app's CSP forbids eval, so tests pass real functions to page.evaluate and reach the modules through window.__load). */
await page.evaluate(() => {
  globalThis.__mods = {};
  globalThis.__load = async (b64) => {
    const bytes = Uint8Array.from(atob(b64), (c) => c.charCodeAt(0));
    const parser = await import('/js/pptx/parser.js');
    const parsed = await parser.parsePptx(bytes);
    const scene = await import('/js/model/scene.js');
    return { parsed, parser, scene };
  };
});
/** Run `fn({ b64, arg })` in the page for a deck file; `fn` calls `await __load(b64)` itself. */
const inPage = (file, fn, arg) => page.evaluate(fn, { b64: b64(file), arg });
const nodeSide = { ok, section, inPage, page, BASE, FX, REAL, realFiles, ROOT, consoleErrors, b64, browser, ctx };

// ======================================================================================================
// STEP 3 — scene model and geometry
// ======================================================================================================
section('scene model: geometry against independently computed corners (p31-geometry)');
{
  const expected = JSON.parse(readFileSync(path.join(FX, 'p31-geometry.expected.json'), 'utf8')).elements;
  const got = await inPage(path.join(FX, 'p31-geometry.pptx'), async ({ b64 }) => {
    const { parsed, scene } = await __load(b64);
    const s = await scene.buildScene(parsed, 0);
    return s.elements.map((e) => ({ name: e.name, id: e.xmlId, sceneId: e.id, kind: e.kind, support: e.support, z: e.z, corners: e.world?.corners || null, parentId: e.parentId, reasons: e.reasons.map((r) => r.text) }));
  });
  const byName = Object.fromEntries(got.map((g) => [g.name, g]));
  let worst = 0;
  for (const [name, exp] of Object.entries(expected)) {
    const g = byName[name];
    ok(!!g, `element ${name} is in the scene`);
    if (!g || !g.corners) continue;
    const d = Math.max(...exp.corners.map((c, i) => Math.abs(c - g.corners[i])));
    worst = Math.max(worst, d);
    ok(d < 1, `${name}: world corners match to < 1 EMU (max error ${d.toFixed(3)})`, d < 1 ? undefined : { exp: exp.corners, got: g.corners });
    ok(String(g.id) === String(exp.id), `${name}: element id is the file's cNvPr id (${exp.id})`);
  }
  ok(worst < 1, `all ${Object.keys(expected).length} elements within 1 EMU (worst ${worst.toFixed(4)})`);
  const order = Object.keys(expected);
  const zs = order.map((n) => byName[n].z);
  ok(zs.every((z, i) => i === 0 || z > zs[i - 1]), 'z-order equals document order (strictly increasing)');
  ok(byName['g06-child1'].parentId === byName['g06-group-scaled'].sceneId, 'group children point at their group');
}

section('scene model: nothing is dropped (every shape element in the XML is in the model) - all fixtures and real decks');
const allDecks = [...readdirSync(FX).filter((f) => f.endsWith('.pptx')).map((f) => path.join(FX, f)), ...realFiles];
{
  let decks = 0; let slides = 0; let elements = 0; let skipped = [];
  for (const f of allDecks) {
    const r = await inPage(f, async ({ b64 }) => {
      let L;
      try { L = await __load(b64); } catch (e) { return { error: e.code || e.message }; }
      const { parsed, scene } = L; const out = [];
      for (let i = 0; i < parsed.slides.length; i++) {
        const sc = await scene.buildScene(parsed, i);
        const ids = sc.elements.map((e) => e.id);
        out.push({ i, n: sc.elements.length, xml: scene.countXmlElements(sc.doc), uniq: new Set(ids).size === ids.length, unplacedOk: sc.unplaced.every((e) => !e.world && e.support === 'opaque'), zOk: sc.elements.every((e, k) => k === 0 || e.z > sc.elements[k - 1].z), allIds: sc.all.every((e) => sc.byId.get(e.id) === e),
          supportOk: sc.elements.every((e) => ['native', 'approximated', 'opaque'].includes(e.support) && (e.support === 'native' || e.reasons.length > 0)), placedOrListed: sc.elements.every((e) => e.world || e.kind === 'group' || sc.unplaced.includes(e) || e.tag === 'AlternateContent') });
      }
      return { slides: out };
    });
    const name = path.basename(f);
    if (r.error) { skipped.push(name); continue; }
    decks++;
    for (const s of r.slides) {
      slides++; elements += s.n;
      ok(s.n === s.xml, `${name} slide ${s.i + 1}: ${s.n} elements in the scene = ${s.xml} in the XML`, s);
      if (!(s.uniq && s.unplacedOk && s.zOk && s.allIds && s.supportOk && s.placedOrListed)) ok(false, `${name} slide ${s.i + 1}: ids unique, z ordered, support+reasons consistent, nothing unplaced-unlisted`, s);
    }
  }
  ok(decks >= 40, `${decks} decks / ${slides} slides / ${elements} elements checked (${skipped.length} invalid files skipped: ${skipped.join(', ')})`);
  ok(realFiles.length >= 8, `${realFiles.length} real-world decks included (tests/real + REAL_PPTX)`);
}

section('scene model: reading never changes the XML, ids are stable across rebuilds');
for (const f of [path.join(FX, 'p31-unsupported.pptx'), path.join(FX, 'p31-template-decor.pptx'), path.join(FX, 'p22-geometry.pptx'), ...realFiles.slice(0, 3)]) {
  const r = await inPage(f, async ({ b64 }) => {
    const { parsed, scene } = await __load(b64);
    const ser = (d) => new XMLSerializer().serializeToString(d);
    const fresh = async (part) => ser(new DOMParser().parseFromString(await parsed.zip.file(part).async('string'), 'application/xml'));
    const res = { same: true, stable: true, parts: 0 };
    for (let i = 0; i < parsed.slides.length; i++) {
      const a = await scene.buildScene(parsed, i); const b = await scene.buildScene(parsed, i);
      res.stable = res.stable && JSON.stringify(a.elements.map((e) => [e.id, e.z, e.world?.corners])) === JSON.stringify(b.elements.map((e) => [e.id, e.z, e.world?.corners]));
      for (const part of new Set([a.part, a.layoutPart, a.masterPart].filter(Boolean))) { res.parts++; res.same = res.same && ser(a.doc.ownerDocument || (part === a.part ? a.doc : (a.layout?.path === part ? a.layout.doc : a.master.doc))) === await fresh(part); }
    }
    return res;
  });
  ok(r.same && r.parts > 0, `${path.basename(f)}: slide/layout/master XML is identical after building scenes (${r.parts} parts)`, r);
  ok(r.stable, `${path.basename(f)}: ids, z-order and geometry are identical on a rebuild`);
}

section('scene model: unsupported and awkward content is kept and labelled (p31-unsupported)');
{
  const els = await inPage(path.join(FX, 'p31-unsupported.pptx'), async ({ b64 }) => {
    const { parsed, scene } = await __load(b64);
    const s = await scene.buildScene(parsed, 0);
    return { unplaced: s.unplaced.map((e) => e.name || e.tag), els: s.elements.map((e) => ({ name: e.name, posSource: e.posSource, tag: e.tag, id: e.id, kind: e.kind, support: e.support, rendered: e.rendered, placed: !!e.world, idIssue: e.idIssue || null, reasons: e.reasons.map((r) => r.text), parentId: e.parentId, imageMissing: !!e.image?.missing })) };
  });
  const by = (n) => els.els.find((e) => e.name === n);
  ok(by('chart-real')?.kind === 'chart' && by('chart-real').support === 'opaque' && by('chart-real').placed, 'chart: kept, positioned, opaque with a reason');
  ok(by('table-merged')?.kind === 'table' && by('table-merged').placed, 'table: kept and positioned');
  ok(by('smartart-synthetic')?.kind === 'diagram' && by('smartart-synthetic').support === 'opaque', 'SmartArt frame: kept, opaque');
  ok(by('ole-synthetic')?.kind === 'ole' && by('ole-synthetic').support === 'opaque', 'OLE frame: kept, opaque');
  const cp = els.els.find((e) => e.tag === 'contentPart');
  ok(cp && cp.kind === 'content' && cp.support === 'opaque' && !cp.placed && els.unplaced.includes('contentPart'), 'contentPart: kept, opaque, listed as unplaced', cp);
  const alt = els.els.find((e) => e.tag === 'AlternateContent');
  ok(alt && alt.kind === 'alt' && alt.support === 'approximated' && /Choice branch/.test(alt.reasons[0]), 'mc:AlternateContent: wrapper kept, branch choice stated', alt);
  ok(by('alt-choice-shape')?.rendered === true && by('alt-choice-shape').id.endsWith('#610'), 'AlternateContent: a Choice made of ordinary shapes is drawn (as PowerPoint does) and keeps the canonical id');
  ok(by('alt-fallback-shape')?.rendered === false && by('alt-fallback-shape').id.endsWith('#610~fallback') && by('alt-fallback-shape').support === 'opaque', 'AlternateContent: the Fallback shape is kept (branch-suffixed id) but marked not drawn');
  ok(by('alt-chartex-fallback')?.rendered === true && by('alt-chartex-fallback').id.endsWith('#611~fallback') && by('alt-chartex-choice')?.rendered === false, 'AlternateContent: a Choice that needs an undrawable graphic frame hands over to the Fallback');
  const unk = els.els.find((e) => e.tag === 'unknownVendorThing');
  ok(unk && unk.kind === 'unknown' && unk.support === 'opaque' && els.unplaced.includes('unknown-element'), 'unknown element: kept, opaque, listed as unplaced');
  ok(by('dup-id-a').id.endsWith('#630') && by('dup-id-b').id.endsWith('#630~2') && by('dup-id-b').idIssue === 'duplicate', 'duplicate ids are made unique and flagged, not merged');
  ok(by('no-id').idIssue === 'missing' && by('no-id').id.includes('@'), 'missing id gets a path-based id and a flag');
  const orphan = by('orphan-placeholder');
  ok(orphan && !orphan.placed && orphan.support === 'opaque' && els.unplaced.includes('orphan-placeholder'), 'placeholder with no resolvable position: kept, listed as unplaced');
  const orphanBody = by('orphan-body-placeholder');
  ok(orphanBody && orphanBody.posSource === 'master' && orphanBody.support !== 'opaque', 'a body placeholder missing from the layout inherits its position from the master body placeholder (matched by type, not by the master date placeholder\'s idx)', orphanBody && orphanBody.posSource);
  ok(by('pic-bad-rel').imageMissing === true, 'picture with a dangling relationship id is kept and flagged missing');
}

section('scene model: master / layout layers, placeholder inheritance, hidden master shapes (p31-template-decor, p31-text)');
{
  const d = await inPage(path.join(FX, 'p31-template-decor.pptx'), async ({ b64 }) => {
    const { parsed, scene } = await __load(b64);
    const out = [];
    for (let i = 0; i < 3; i++) { const s = await scene.buildScene(parsed, i); out.push({ layers: s.layers.map((l) => ({ o: l.origin, names: l.elements.map((e) => e.name) })), bg: s.background.source, hidden: s.masterShapesHidden }); }
    return out;
  });
  ok(d[0].layers.map((l) => l.o).join() === 'master,layout,slide', 'slide 1: layers are master, layout, slide (bottom to top)');
  ok(d[0].layers[0].names.includes('master-bar') && d[0].layers[1].names.includes('layout-badge'), 'master and layout decoration are present, placeholders are not');
  ok(!d[0].layers[0].names.some((n) => /Title|Text Placeholder/.test(n)) && !d[0].layers[1].names.some((n) => /Title|Content/.test(n)), 'template placeholders are never layer content');
  ok(d[1].layers.map((l) => l.o).join() === 'slide' && d[1].hidden, 'slide 2: showMasterSp="0" hides master and layout shapes');
  ok(d[0].bg === 'master' && d[2].bg === 'slide', 'background source: master gradient on slide 1, the slide\'s own picture on slide 3');
  const t = await inPage(path.join(FX, 'p31-text.pptx'), async ({ b64 }) => {
    const { parsed, scene } = await __load(b64);
    const s = await scene.buildScene(parsed, 2);
    return s.elements.map((e) => ({ ph: e.ph && e.ph.type, src: e.posSource, w: e.world && Math.round(e.world.w), h: e.world && Math.round(e.world.h), x: e.world && Math.round(e.world.cx - e.world.w / 2) }));
  });
  ok(t.length === 2 && t.every((e) => e.src === 'layout' || e.src === 'master') && t.every((e) => e.w > 0), 'placeholders inherit their position from the layout / master', t);
}




// ==== Step 4: style resolution ====
section('style resolution: colour engine (values PowerPoint shows for its own theme variants)');
{
  const r = await page.evaluate(async () => {
    const { colorEl } = await import('/js/model/resolve.js');
    const mk = (x) => new DOMParser().parseFromString(`<a:solidFill xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">${x}</a:solidFill>`, 'application/xml').documentElement.firstElementChild;
    const cx = { theme: { colors: { accent1: '4472C4', dk1: '000000', lt1: 'FFFFFF', tx1: '000000' } }, clrMap: { tx1: 'dk1', bg1: 'lt1' } };
    const h = (x, ph) => { const c = colorEl(mk(x), cx, ph); return c && c.hex + (c.alpha !== 1 ? '@' + c.alpha : ''); };
    return {
      darker25: h('<a:schemeClr val="accent1"><a:lumMod val="75000"/></a:schemeClr>'),
      lighter40: h('<a:schemeClr val="accent1"><a:lumMod val="60000"/><a:lumOff val="40000"/></a:schemeClr>'),
      lighter80: h('<a:schemeClr val="accent1"><a:lumMod val="20000"/><a:lumOff val="80000"/></a:schemeClr>'),
      mapped: h('<a:schemeClr val="tx1"/>'),
      alpha: h('<a:srgbClr val="FF0000"><a:alpha val="50000"/></a:srgbClr>'),
      phClr: h('<a:schemeClr val="phClr"><a:shade val="50000"/></a:schemeClr>', { hex: 'FFFFFF', alpha: 1 }),
      prst: h('<a:prstClr val="red"/>'),
      sys: h('<a:sysClr val="windowText" lastClr="000000"/>'),
      tint: h('<a:srgbClr val="000000"><a:tint val="50000"/></a:srgbClr>'),
      unknownMod: colorEl(mk('<a:srgbClr val="808080"><a:comp/></a:srgbClr>'), cx).approx,
      noPh: colorEl(mk('<a:schemeClr val="phClr"/>'), cx),
    };
  });
  ok(r.darker25 === '2F5597', 'accent1 "darker 25%" = 2F5597', r.darker25);
  ok(r.lighter40 === '8FAADC', 'accent1 "lighter 40%" = 8FAADC', r.lighter40);
  ok(r.lighter80 === 'DAE3F3', 'accent1 "lighter 80%" = DAE3F3', r.lighter80);
  ok(r.mapped === '000000', 'scheme colour goes through the colour map (tx1 -> dk1)');
  ok(r.alpha === 'FF0000@0.5', 'alpha is kept separately from the colour', r.alpha);
  ok(/^BC/.test(r.phClr), 'phClr takes the style reference colour and applies its own transforms (white, shade 50% in linear light = BCBCBC)', r.phClr);
  ok(r.prst === 'FF0000' && r.sys === '000000', 'preset and system colours');
  ok(/^BC/.test(r.tint), 'tint is applied in linear light (black at 50% tint = BCBCBC, not 808080)', r.tint);
  ok(r.unknownMod && r.unknownMod.includes('comp') && r.noPh === null, 'unsupported colour modifiers are reported, phClr without a reference is null');
}

section('style resolution: all fixtures and real decks');
const resolveAll = [path.join(FX, 'p31-geometry.pptx'), path.join(FX, 'p31-shapes.pptx'), path.join(FX, 'p31-text.pptx'), path.join(FX, 'p31-arabic.pptx'), path.join(FX, 'p31-media.pptx'), path.join(FX, 'p31-unsupported.pptx'), path.join(FX, 'p31-template-decor.pptx'), ...realFiles];
for (const f of resolveAll) {
  const r = await inPage(f, async ({ b64 }) => {
    const { parsed, scene } = await __load(b64);
    const rs = await import('/js/model/resolve.js');
    const ser = (d) => new XMLSerializer().serializeToString(d);
    const res = { slides: parsed.slides.length, err: null, mutated: 0, textLost: 0, textChecked: 0, unresolved: 0, idem: true, bgMissing: 0, badSupport: 0 };
    try {
      for (let i = 0; i < parsed.slides.length; i++) {
        const s = await scene.buildScene(parsed, i);
        const before = [...new Set(s.layers.map((l) => l.part))].map((p) => [p, ser(s.layers.find((l) => l.part === p).elements[0]?.el.ownerDocument || s.doc)]);
        await rs.resolveScene(s);
        const snap = JSON.stringify(s.all.map((e) => [e.id, e.support, e.reasons.length]));
        await rs.resolveScene(s);
        if (JSON.stringify(s.all.map((e) => [e.id, e.support, e.reasons.length])) !== snap) res.idem = false;
        for (const [p, x] of before) { const d = s.layers.find((l) => l.part === p).elements[0]?.el.ownerDocument; if (d && ser(d) !== x) res.mutated++; }
        if (!s.backgroundFill || !s.backgroundFill.type) res.bgMissing++;
        for (const e of s.all) {
          if (!['native', 'approximated', 'opaque'].includes(e.support)) res.badSupport++;
          if (e.tag === 'sp' && e.el.getElementsByTagNameNS('*', 'txBody').length && !e.text) res.unresolved++;
          if (e.text) {
            const want = [...e.el.getElementsByTagNameNS('*', 'p')].map((p) => [...p.children].filter((c) => ['r', 'fld'].includes(c.localName)).map((r) => [...r.children].filter((c) => c.localName === 't').map((t) => t.textContent).join('')).join(''));
            const got = e.text.paras.map((p) => p.runs.filter((r) => !r.br).map((r) => r.text).join(''));
            res.textChecked++;
            if (want.join('\n') !== got.join('\n')) res.textLost++;
          }
          if (e.tag === 'sp' && !e.style) res.unresolved++;
        }
      }
    } catch (x) { res.err = String(x && x.stack || x).slice(0, 400); }
    return res;
  });
  const n = path.basename(f);
  ok(!r.err, `${n}: resolves without error`, r.err);
  ok(r.mutated === 0, `${n}: resolving never touches the XML`, r.mutated);
  ok(r.textLost === 0, `${n}: text of every shape is preserved exactly (${r.textChecked} text bodies)`, r.textLost);
  ok(r.unresolved === 0 && r.bgMissing === 0 && r.badSupport === 0 && r.idem, `${n}: every shape has a style, every slide a background, support levels valid, idempotent`, r);
}

section('style resolution: fixture-specific values');
{
  const sh = await inPage(path.join(FX, 'p31-shapes.pptx'), async ({ b64 }) => {
    const { parsed, scene } = await __load(b64); const rs = await import('/js/model/resolve.js');
    const s = await scene.buildScene(parsed, 1); await rs.resolveScene(s);
    const g = (n) => { const e = s.all.find((x) => x.id.endsWith('#' + n)); return { fill: e.style?.fill, line: e.style?.line, shadow: e.style?.shadow, glow: e.style?.glow, support: e.support, reasons: e.reasons.map((r) => r.text) }; };
    return { grad: g(202), path: g(203), patt: g(204), dash: g(211), arrow: g(220), shadow: g(221), glow: g(222), themed: g(224) };
  });
  ok(sh.grad.fill.type === 'gradient' && sh.grad.fill.kind === 'linear' && sh.grad.fill.angle === 45 && sh.grad.fill.stops.length >= 2, 'linear gradient: angle and stops');
  ok(sh.path.support === 'approximated' && /gradient/.test(sh.path.reasons.join()), 'path gradient is approximated with a reason');
  ok(sh.patt.support === 'approximated' && sh.patt.fill.type === 'pattern', 'pattern fill is approximated with a reason');
  ok(sh.dash.line.dash === 'dash' && sh.dash.line.w === 38100 && sh.dash.line.cap === 'round', 'outline width, dash and cap');
  ok(sh.arrow.line.head?.type === 'triangle' && sh.arrow.line.tail?.type === 'arrow' && sh.arrow.line.tail.w === 'lg', 'line arrowheads (head / tail, size)');
  ok(sh.shadow.shadow && sh.shadow.shadow.blur > 0 && sh.shadow.support === 'native', 'outer shadow resolved (offset, blur, colour)');
  ok(sh.glow.support === 'approximated' && /glow/.test(sh.glow.reasons.join()) && sh.glow.glow?.rad > 0, 'glow is drawn as a halo and reported as approximated, never silently dropped');
  ok(sh.themed.fill.hex === 'C0504D' && sh.themed.line.w === 25400, 'fill and outline from the theme style matrix (fillRef / lnRef + phClr)', sh.themed);

  const tx = await inPage(path.join(FX, 'p31-text.pptx'), async ({ b64 }) => {
    const { parsed, scene } = await __load(b64); const rs = await import('/js/model/resolve.js');
    const s = await scene.buildScene(parsed, 0); await rs.resolveScene(s);
    return s.elements.filter((e) => e.text).map((e) => e.text.paras);
  });
  const first = tx[0][0].runs;
  ok(first[0].b === true && first[1].i === true && first[2].u === true && first[3].hex === 'C00000' && first[4].font === 'Georgia', 'runs: bold, italic, underline, colour, own font');
  ok(first.some((r) => r.baseline > 0) && first.some((r) => r.baseline < 0), 'superscript and subscript baseline offsets');
  ok(tx[0][1].algn === 'center' && tx[0][1].runs[0].sz === 14 && tx[0][1].runs[0].strike === true, 'paragraph alignment, size and strikethrough');
  const list = tx[1];
  ok(list[0].bullet?.char === '•' && list[0].marL === 342900 && list[0].indent === -342900, 'bullet character with hanging indent');
  ok(list[1].lvl === 1 && list[1].bullet?.char === '–', 'second-level bullet from the paragraph level');
  ok(list[2].bullet?.label === '1.' && list[3].bullet?.label === '2.', 'automatic numbering counts up');
  ok(list[4].bullet === null, 'buNone suppresses the bullet');
  ok(tx[2][0].lnSpc?.pct === 1.5 && tx[2][0].spcBef?.pts === 6 && tx[2][1].lnSpc?.pts === 30 && tx[2][1].spcAft?.pts === 12, 'line spacing (percent and exact points), space before / after');

  const ar = await inPage(path.join(FX, 'p31-arabic.pptx'), async ({ b64 }) => {
    const { parsed, scene } = await __load(b64); const rs = await import('/js/model/resolve.js');
    const s = await scene.buildScene(parsed, 1); await rs.resolveScene(s);
    const t = (n) => s.all.find((x) => x.id.endsWith('#' + n)).text.paras;
    return { a400: t(400), a402: t(402), a404: t(404), a405: t(405), a407: t(407), table: s.all.find((x) => x.id.endsWith('#408')).table };
  });
  ok(ar.a400[0].rtl && ar.a400[0].algn === 'right' && ar.a400[0].runs[0].rtl, 'Arabic paragraph: rtl, right-aligned, run flagged rtl');
  ok(ar.a402[0].runs.length >= 3 && ar.a402[0].runs.some((r) => !r.rtl && /Nabtex/.test(r.text)) && ar.a402[0].runs.some((r) => r.rtl), 'mixed Arabic / Latin run is split by script so each word gets its own font');
  ok(ar.a402[0].runs.find((r) => /Nabtex/.test(r.text)).font === 'Calibri' && ar.a402[0].runs.find((r) => r.rtl).font === 'Arial', 'Latin and Arabic segments resolve their fonts independently');
  ok(!ar.a404[0].rtl && ar.a404[0].algn === 'left', 'English left-to-right shape stays ltr');
  ok(ar.a405[0].algn === 'justify' && ar.a405[0].rtl, 'justified Arabic paragraph');
  ok(ar.table.cols.length === 2 && ar.table.rows.length === 2 && ar.table.rows[0].cells[0].fill?.hex === '4F81BD' && ar.table.rows[0].cells[0].text[0].runs[0].b === true, 'table: grid, header-row style fill and bold text from the table style');

  const un = await inPage(path.join(FX, 'p31-unsupported.pptx'), async ({ b64 }) => {
    const { parsed, scene } = await __load(b64); const rs = await import('/js/model/resolve.js');
    const s = await scene.buildScene(parsed, 0); await rs.resolveScene(s);
    return s.all.map((e) => ({ id: e.id, kind: e.kind, support: e.support, hasStyle: !!e.style, rendered: e.rendered }));
  });
  const by = (suffix) => un.find((e) => e.id.endsWith(suffix));
  ok(by('#2').support === 'opaque' && !by('#2').hasStyle && by('#600').support === 'opaque' && by('#601').support === 'opaque', 'chart / SmartArt / OLE stay opaque after resolution');
  ok(by('#610').support === 'native' && by('#610~fallback').support === 'opaque' && by('#611~fallback').support === 'native', 'AlternateContent: undrawn branch stays opaque, drawn branch native');
  const md = await inPage(path.join(FX, 'p31-media.pptx'), async ({ b64 }) => {
    const { parsed, scene } = await __load(b64); const rs = await import('/js/model/resolve.js');
    const s = await scene.buildScene(parsed, 0); await rs.resolveScene(s);
    return s.all.map((e) => ({ id: e.id, support: e.support, crop: e.image && e.image.crop, reasons: e.reasons.map((r) => r.text) }));
  });
  const mby = (n) => md.find((e) => e.id.endsWith('#' + n));
  ok(mby(3).crop && mby(3).crop.l === 25000, 'picture crop (srcRect) resolved');
  ok(mby(501).support === 'approximated' && /EMF/.test(mby(501).reasons.join()), 'EMF picture is reported as not displayable');
  ok(mby(502).support === 'approximated' && /linked/.test(mby(502).reasons.join()), 'linked (external) picture is reported, nothing is fetched');
  ok(mby(503).support === 'approximated' && /duotone/.test(mby(503).reasons.join()), 'picture effects that are not applied are reported');
}


// ==== Step 5: renderer ====
import { mkdirSync, mkdtempSync, rmSync } from 'node:fs';
import os from 'node:os';
const KEEP_SHOTS = !!process.env.SHOTS;
const SHOTS = process.env.SHOTS || mkdtempSync(path.join(os.tmpdir(), 'sf-shots-')); mkdirSync(SHOTS, { recursive: true });
await page.evaluate(() => {
  const h = document.createElement('div'); h.id = 'sf-host'; document.body.append(h);
  const st = h.style; st.position = 'fixed'; st.left = '0'; st.top = '0'; st.width = '960px'; st.zIndex = '99999'; st.background = '#888';
  globalThis.__render = async (b64, idx, width) => {
    const { parsed, scene } = await globalThis.__load(b64);
    const rn = await import('/js/render/render.js');
    const sc = await scene.buildScene(parsed, idx);
    const host = document.getElementById('sf-host'); host.hidden = false;
    const hnd = await rn.renderSlide(sc, host, { width });
    return { hnd, sc, parsed };
  };
});
async function shoot(file, idx, name, width = 960) {
  const rep = await page.evaluate(async ({ b64, idx, width }) => {
    const { hnd, sc } = await globalThis.__render(b64, idx, width);
    const r = hnd.report();
    return { counts: r.counts, issues: r.issues.map((i) => i.id + ' ' + i.support + ' ' + i.reasons.join('; ')), nodes: hnd.root.querySelectorAll('.sf-el').length, canvases: document.querySelectorAll('#sf-host canvas').length, total: r.total, bg: r.background, w: hnd.root.getBoundingClientRect().width, h: hnd.root.getBoundingClientRect().height };
  }, { b64: b64(file), idx, width });
  await page.locator('#sf-host .sf-slide').screenshot({ path: path.join(SHOTS, name + '.png') });
  return rep;
}
section('renderer: every fixture and real deck draws without errors');
const renderDecks = [...['p31-geometry', 'p31-shapes', 'p31-text', 'p31-arabic', 'p31-media', 'p31-unsupported', 'p31-template-decor'].map((n) => path.join(FX, n + '.pptx')), ...realFiles];
for (const f of renderDecks) {
  const n = await inPage(f, async ({ b64 }) => { const { parsed } = await __load(b64); return parsed.slides.length; });
  let bad = 0; let nodes = 0; let total = 0; let canv = 0; let err = null; let size = true;
  try {
    for (let i = 0; i < n; i++) {
      const r = await shoot(f, i, path.basename(f, '.pptx') + '-' + (i + 1));
      nodes += r.nodes; total += r.total; canv += r.canvases;
      const unplaced = r.total - r.nodes; if (unplaced < 0) bad++;
      if (Math.abs(r.w - 960) > 1) size = false;
    }
  } catch (x) { err = String(x).slice(0, 300); }
  ok(!err, `${path.basename(f)}: ${n} slide(s) render`, err);
  ok(bad === 0 && canv === 0 && size, `${path.basename(f)}: no canvas / raster, slide fills the requested width`, { bad, canv, size });
}


section('renderer: structure of the drawn DOM (ids, z-order, text, geometry, no raster, XML untouched)');
for (const f of renderDecks) {
  const n = await inPage(f, async ({ b64 }) => { const { parsed } = await __load(b64); return parsed.slides.length; });
  const agg = { slides: 0, idMismatch: 0, zBad: 0, textBad: 0, textChecked: 0, geomBad: 0, geomChecked: 0, raster: 0, mutated: 0, hiddenShown: 0, supportMismatch: 0, imgTags: 0 };
  for (let i = 0; i < n; i++) {
    const r = await page.evaluate(async ({ b64, i }) => {
      const { parsed, scene } = await globalThis.__load(b64);
      const sc = await scene.buildScene(parsed, i);
      const ser = (d) => new XMLSerializer().serializeToString(d);
      const docs = [...new Set(sc.all.map((e) => e.el.ownerDocument))]; const before = docs.map(ser);
      const rn = await import('/js/render/render.js');
      const host = document.getElementById('sf-host');
      const hnd = await rn.renderSlide(sc, host, { width: 960 });
      const out = { idMismatch: 0, zBad: 0, textBad: 0, textChecked: 0, geomBad: 0, geomChecked: 0, raster: 0, mutated: 0, hiddenShown: 0, supportMismatch: 0, imgTags: 0 };
      const want = sc.all.filter((e) => e.rendered && e.world && e.kind !== 'alt').map((e) => e.id);
      const got = [...hnd.root.querySelectorAll('.sf-el')].map((n) => n.dataset.id);
      if (want.join('|') !== got.join('|')) out.idMismatch++; // same ids, same order = document / z-order
      const zs = [...hnd.root.querySelectorAll('.sf-el')].map((n) => Number(n.dataset.z)); if (zs.some((z, k) => k && z < zs[k - 1])) out.zBad++;
      out.raster = host.querySelectorAll('canvas, video').length; out.imgTags = host.querySelectorAll('img').length;
      const k = 960 / sc.width;
      for (const e of sc.all) {
        const node = hnd.nodes.get(e.id); if (!node) continue;
        if (node.dataset.support !== e.support) out.supportMismatch++;
        if (e.hidden && getComputedStyle(node).display !== 'none') out.hiddenShown++;
        if (e.text && e.kind === 'shape' && !e.hidden) {
          const model = e.text.paras.map((p) => p.runs.filter((r) => !r.br).map((r) => r.text).join(''));
          const dom = [...node.querySelectorAll('.sf-p')].map((p) => [...p.childNodes].filter((c) => !(c.classList && c.classList.contains('sf-bu'))).map((c) => c.textContent).join('').replace(/​/g, ''));
          out.textChecked++; if (model.join('\n') !== dom.join('\n')) out.textBad++;
        }
        if (e.world && !e.world.rot && !e.hidden && e.world.w > 0 && e.world.h > 0) {
          const rc = node.getBoundingClientRect(); const rr = hnd.root.getBoundingClientRect();
          const ex = (e.world.cx - e.world.w / 2) * k; const ey = (e.world.cy - e.world.h / 2) * k;
          out.geomChecked++;
          if (Math.abs(rc.left - rr.left - ex) > 0.6 || Math.abs(rc.top - rr.top - ey) > 0.6 || Math.abs(rc.width - e.world.w * k) > 0.6 || Math.abs(rc.height - e.world.h * k) > 0.6) out.geomBad++;
        }
      }
      docs.forEach((d, j) => { if (ser(d) !== before[j]) out.mutated++; });
      return out;
    }, { b64: b64(f), i });
    agg.slides++; for (const k of Object.keys(r)) agg[k] += r[k];
  }
  const nm = path.basename(f);
  ok(agg.idMismatch === 0 && agg.zBad === 0, `${nm}: every placed element is drawn once, with its stable id, in document (z) order`, agg);
  ok(agg.textBad === 0, `${nm}: DOM text equals the model text (${agg.textChecked} text shapes) and is real selectable text`, agg);
  ok(agg.geomBad === 0, `${nm}: unrotated elements sit at their exact position and size (${agg.geomChecked} checked, 0.6 px)`, agg);
  ok(agg.raster === 0 && agg.imgTags === 0 && agg.mutated === 0 && agg.hiddenShown === 0 && agg.supportMismatch === 0, `${nm}: no canvas / <img> raster, XML untouched, hidden elements hidden, support labels match`, agg);
}

section('renderer: native content (no pixels)');
{
  const r = await page.evaluate(async ({ b64 }) => {
    const { hnd } = await globalThis.__render(b64, 1, 960);
    const root = hnd.root;
    const cs = (sel) => root.querySelectorAll(sel).length;
    return { svg: cs('svg.sf-geo'), paths: cs('svg.sf-geo path'), text: cs('.sf-text'), flipped: [...root.querySelectorAll('svg.sf-geo')].filter((s) => s.style.transform === 'scaleX(-1)').length, markers: cs('marker'), gradients: cs('linearGradient, radialGradient'),
      shadows: [...root.querySelectorAll('svg.sf-geo')].filter((s) => /drop-shadow/.test(s.style.filter)).length, inlineStyleAttr: [...root.querySelectorAll('[style]')].length };
  }, { b64: b64(path.join(FX, 'p31-shapes.pptx')) });
  ok(r.svg > 10 && r.paths >= r.svg, 'shapes are SVG paths (vector), not images', r);
  ok(r.markers >= 2 && r.gradients >= 2 && r.shadows >= 1, 'arrow heads are SVG markers, gradients are SVG gradients, shadows are CSS filters', r);
  ok(r.flipped >= 1, 'flipped shapes use a mirrored vector, text is left unmirrored');
  const pic = await page.evaluate(async ({ b64 }) => { const { hnd } = await globalThis.__render(b64, 0, 960); return { images: hnd.root.querySelectorAll('svg image').length, clip: hnd.root.querySelectorAll('clipPath').length, role: hnd.root.querySelectorAll('[role=img]').length }; }, { b64: b64(path.join(FX, 'p31-media.pptx')) });
  ok(pic.images >= 5 && pic.clip >= 5, 'pictures are vector-clipped <image> elements with crop and geometry', pic);
}


// ==== Step 6: fonts ====
import crypto from 'node:crypto';
section('fonts: self-hosted files, licences, CSP');
const FONTDIR = path.join(ROOT, 'fonts');
const fontFiles = readdirSync(FONTDIR).filter((f) => f.endsWith('.ttf'));
const readme = readFileSync(path.join(FONTDIR, 'README.md'), 'utf8');
const fontCss = readFileSync(path.join(ROOT, 'css/fonts.css'), 'utf8');
ok(fontFiles.length === 20, 'twenty font files are shipped', fontFiles.length);
ok(fontFiles.every((f) => { const h = crypto.createHash('sha256').update(readFileSync(path.join(FONTDIR, f))).digest('hex'); return readme.includes(`${h}  ${f}`); }), 'every file is byte-identical to the checksum recorded in fonts/README.md (unmodified originals)');
ok(fontFiles.every((f) => fontCss.includes(`../fonts/${f}`)), 'every shipped file is declared in css/fonts.css');
ok(fontFiles.every((f) => /^[A-Za-z]+(-[A-Za-z]+)?\.ttf$/.test(f)) && !fontFiles.some((f) => /calibri|arial|times|cambria|segoe|tahoma|verdana|georgia|consola|courier new/i.test(f)), 'no commercial fonts are shipped (only Carlito, Caladea, Liberation, DejaVu)');
const lic = readdirSync(path.join(FONTDIR, 'licenses'));
ok(['OFL-Carlito.txt', 'OFL-Liberation.txt', 'Apache-2.0-Caladea.txt', 'DejaVu-Bitstream-Vera.txt'].every((l) => lic.includes(l)), 'licence text ships with every font family');
ok(readFileSync(path.join(FONTDIR, 'licenses/OFL-Carlito.txt'), 'utf8').includes('SIL OPEN FONT LICENSE') && readFileSync(path.join(FONTDIR, 'licenses/Apache-2.0-Caladea.txt'), 'utf8').includes('Apache License'), 'licence files contain the full licence text');

const fontReqs = [];
page.on('request', (r) => { if (r.resourceType() === 'font') fontReqs.push(r.url()); });
const cdp = await ctx.newCDPSession(page);
await cdp.send('DOM.enable'); await cdp.send('CSS.enable');
const platformFonts = async (html, css) => {
  await page.evaluate(({ html, css }) => { document.getElementById('ftest')?.remove(); const d = document.createElement('div'); d.id = 'ftest'; d.style.cssText = 'position:fixed;left:0;top:0;z-index:100000;background:#fff;font-size:28px;' + css; d.innerHTML = html; document.body.append(d); }, { html, css });
  await page.evaluate(async () => { await document.fonts.ready; });
  await page.evaluate(() => document.fonts.load('400 28px "SF Carlito"').catch(() => 0));
  const doc = await cdp.send('DOM.getDocument', { depth: 0 });
  const q = await cdp.send('DOM.querySelector', { nodeId: doc.root.nodeId, selector: '#ftest span' });
  const r = await cdp.send('CSS.getPlatformFontsForNode', { nodeId: q.nodeId });
  return r.fonts;
};
section('fonts: each self-hosted face loads and is what the page really draws when the original is missing');
{
  const faces = [['SF Carlito', 'Calibri'], ['SF Caladea', 'Cambria'], ['SF Liberation Sans', 'Arial'], ['SF Liberation Serif', 'Times New Roman'], ['SF Liberation Mono', 'Courier New'], ['SF DejaVu Sans', 'Arabic']];
  for (const [fam, stands] of faces) {
    const loaded = await page.evaluate(async (fam) => { const out = []; for (const w of [400, 700]) { const r = await document.fonts.load(`${w} 24px "${fam}"`, 'abc'); out.push(r.length ? r.every((x) => x.status === 'loaded') : false); } return out; }, fam);
    ok(loaded.every(Boolean), `${fam} (for ${stands}): regular and bold load from the app origin`, loaded);
    const used = await platformFonts('<span>Hamburgefonstiv 0123</span>', `font-family:"No Such Font ${fam}","${fam}",serif`);
    ok(used.length === 1 && used[0].isCustomFont === true, `${fam}: drawn by the self-hosted file when the deck's font is missing`, used);
  }
  ok(fontReqs.length > 0 && fontReqs.every((u) => u.startsWith(BASE + '/fonts/')), 'every font request goes to the app\'s own /fonts (nothing third-party)', fontReqs.filter((u) => !u.startsWith(BASE + '/fonts/')));
  const orig = await platformFonts('<span>Hamburgefonstiv 0123</span>', 'font-family:"DejaVu Serif","SF Carlito",serif');
  ok(orig.length === 1 && orig[0].isCustomFont === false && /DejaVu Serif/.test(orig[0].familyName), 'the deck\'s own font is used first when it is installed (never replaced by a substitute)', orig);
  const stackOk = await page.evaluate(async () => { const m = await import('/js/render/fonts.js'); return { calibri: m.fontStack('Calibri'), ar: m.fontStack('Arial', true), mono: m.fontStack('Courier New'), unknown: m.fontStack('Brush Script X'), theme: m.fontStack('+mn-lt') }; });
  ok(/^"Calibri", "SF Carlito", sans-serif$/.test(stackOk.calibri) && /^"Arial", "SF Liberation Sans", "Noto Naskh Arabic", "Noto Sans Arabic", "SF DejaVu Sans", sans-serif$/.test(stackOk.ar), 'stacks: original font, then its substitute, then (Arabic) an Arabic-capable face, then a generic family', stackOk);
  ok(/monospace$/.test(stackOk.mono) && /^"Brush Script X", sans-serif$/.test(stackOk.unknown), 'unknown fonts are kept first and get a generic fallback only', stackOk);
}

section('fonts: Arabic and English slides');
{
  const arabicRun = async (file, idx, pick) => {
    const rep = await page.evaluate(async ({ b64, idx }) => {
      const { hnd } = await globalThis.__render(b64, idx, 960);
      document.querySelectorAll('.sf-r').forEach((n, k) => { n.dataset.k = String(k); });
      const spans = [...hnd.root.querySelectorAll('.sf-r')].map((n) => ({ k: n.dataset.k, text: n.textContent, ar: /[؀-ۿ]/.test(n.textContent), lang: n.lang, family: n.style.fontFamily }));
      return { spans, fonts: hnd.report().fonts };
    }, { b64: b64(file), idx });
    await page.evaluate(() => new Promise((r) => requestAnimationFrame(() => requestAnimationFrame(r)))); // let one frame lay the text out so font usage is recorded
    const doc = await cdp.send('DOM.getDocument', { depth: 0 });
    const out = [];
    for (const sp of rep.spans.filter(pick)) {
      const qn = await cdp.send('DOM.querySelector', { nodeId: doc.root.nodeId, selector: `.sf-r[data-k="${sp.k}"]` });
      const pf = await cdp.send('CSS.getPlatformFontsForNode', { nodeId: qn.nodeId });
      out.push({ ...sp, platform: pf.fonts });
    }
    return { out, fonts: rep.fonts };
  };
  // Arabic letters must be drawn by an Arabic-capable face; spaces / digits / Latin may come from the deck's own (Latin) font.
  const arabicOk = (s) => { const letters = [...s.text].filter((c) => /[\u0600-\u06FF]/.test(c) && !/[\u0660-\u0669\u066A-\u066D]/.test(c)).length; const arabicFace = s.platform.filter((p) => p.isCustomFont || /Noto|Naskh/i.test(p.familyName)).reduce((n, p) => n + p.glyphCount, 0); const other = s.platform.filter((p) => !(p.isCustomFont || /Noto|Naskh/i.test(p.familyName))).reduce((n, p) => n + p.glyphCount, 0); const neutral = [...s.text].filter((c) => !/[\u0600-\u06FF]/.test(c)).length; return letters === 0 || (arabicFace >= letters * 0.6 && other <= neutral + 1); };
  const ar = await arabicRun(path.join(FX, 'p31-arabic.pptx'), 1, (s) => s.text.trim());
  const arSpans = ar.out.filter((s) => s.ar); const enSpans = ar.out.filter((s) => !s.ar && /[A-Za-z]/.test(s.text));
  ok(arSpans.length >= 6 && arSpans.every((s) => s.lang === 'ar'), 'Arabic runs are tagged lang="ar"', arSpans.length);
  ok(arSpans.every(arabicOk), 'every Arabic letter is drawn by an Arabic-capable face (self-hosted DejaVu Sans or an installed Noto); only spaces / digits come from the Latin font', arSpans.filter((s) => !arabicOk(s)).slice(0, 2));
  ok(arSpans.every((s) => s.platform.reduce((n, p) => n + p.glyphCount, 0) >= [...s.text].length * 0.6), 'Arabic glyph count is consistent with the text (shaping / ligatures, nothing dropped)');
  ok(enSpans.length >= 2 && enSpans.every((s) => s.platform.length >= 1 && s.platform.every((p) => !/DejaVu/.test(p.familyName) || false)), 'English runs on the same slide use the Latin font, not the Arabic one', enSpans.map((s) => [s.text, s.platform]).slice(0, 2));
  ok(ar.fonts.some((f) => f.arabic), 'the font report flags the fonts used for Arabic text', ar.fonts);
  // a real PowerPoint-authored Arabic deck
  const realAr = realFiles.find((f) => /arabic/i.test(f));
  if (realAr) {
    const r = await arabicRun(realAr, 0, (s) => s.text.trim());
    ok(r.out.some((s) => s.ar) && r.out.filter((s) => s.ar).every(arabicOk), `${path.basename(realAr)}: Arabic text of the real deck is drawn by an Arabic-capable face`, r.out.slice(0, 2));
  }
  const en = await arabicRun(path.join(FX, 'p31-text.pptx'), 0, (s) => s.text.trim());
  ok(en.out.length > 5 && en.out.every((s) => s.platform.length >= 1 && s.platform.every((p) => p.glyphCount > 0)), 'English text renders (every run has a real font and glyphs)');
  ok(en.out.some((s) => /Georgia/.test(s.family)) && en.fonts.some((f) => f.name === 'Georgia'), 'a run that asks for Georgia keeps Georgia first in its stack and is listed in the font report', en.fonts);
}


// ==== Step 7: the Rendered tab in the real UI ====
section('Rendered tab: real UI flow (upload, preview, tabs)');
{
  const ui = await ctx.newPage();
  ui.on('pageerror', (e) => consoleErrors.push('ui pageerror: ' + e.message));
  ui.on('console', (m) => { if (m.type() === 'error' && !(/Applying inline style/.test(m.text()) && m.location().url.endsWith('/js/pptx/parser.js'))) consoleErrors.push('ui: ' + m.text()); });
  await ui.setViewportSize({ width: 1280, height: 900 });
  await ui.goto(BASE + '/index.html'); await ui.waitForSelector('.shell');
  await ui.click('text=Templates >> nth=0'); await ui.waitForSelector('.dropzone');
  const upload = async (file, name) => {
    const [fc] = await Promise.all([ui.waitForEvent('filechooser'), ui.click('.dropzone')]); await fc.setFiles(path.join(FX, file));
    await ui.waitForSelector('.modal #tpl-name'); await ui.fill('#tpl-name', name); await ui.click('.modal >> text=Save to library');
    await ui.waitForFunction((n) => [...document.querySelectorAll('.tpl h3')].some((x) => x.textContent === n), name);
  };
  await upload('p31-arabic.pptx', 'Arabic deck');
  await upload('p31-unsupported.pptx', 'Unsupported deck');
  await ui.click('.tpl:has-text("Arabic deck") >> text=Preview'); await ui.waitForSelector('.film-item');
  const tabNames = await ui.locator('.tabs [role=tab]').allTextContents();
  ok(tabNames.join('|') === 'Slides|Design|Rendered', 'the preview has Slides, Design and Rendered tabs (existing two unchanged)', tabNames);
  ok((await ui.getAttribute('role=tab[name="Slides"]', 'aria-selected')) === 'true', 'Slides stays the default tab');
  ok((await ui.locator('.rendered-host .sf-slide').count()) === 0, 'nothing is rendered (and no font / CSS downloaded) until the Rendered tab is opened');
  const slidesBefore = await ui.locator('.slide .shp').count();
  await ui.click('role=tab[name="Rendered"]');
  await ui.waitForSelector('.rendered-host .sf-slide .sf-el');
  const info = await ui.evaluate(() => { const r = document.querySelector('.rendered-host .sf-slide'); const f = document.querySelector('.sf-frame'); return { w: r.getBoundingClientRect().width, fw: f.clientWidth, els: r.querySelectorAll('.sf-el').length, text: r.querySelectorAll('.sf-text').length, canvas: document.querySelectorAll('canvas').length, editable: [...document.querySelectorAll('.rendered [contenteditable], .rendered input, .rendered textarea, .rendered button[type=submit]')].filter((n) => !n.closest('.sf-style-bar')).length, aria: r.getAttribute('aria-label') }; });
  ok(info.els >= 2 && info.text >= 1 && info.canvas === 0, 'Rendered tab draws the slide as native elements', info);
  ok(Math.abs(info.w - info.fw) <= 1, 'the rendered slide fills its frame', info);
  ok(info.editable === 0, 'the tab is read-only: no inputs, no contenteditable, no save controls', info);
  ok(/rendered from the file \(read-only\)/.test(info.aria), 'the slide is labelled for assistive technology');
  const det = await ui.textContent('.rendered .details');
  ok(/Support on this slide/.test(det) && /Native/.test(det) && /Fonts/.test(det), 'support summary and font report are shown');
  ok((await ui.textContent('.rendered .stage-wrap .fine')).includes('drawn from the file'), 'the tab says what it is showing');
  if (SHOTS) await ui.screenshot({ path: path.join(SHOTS, 'ui-rendered-arabic.png') });

  // navigation is shared with the Slides tab
  await ui.click('.rendered .film-item >> nth=1');
  await ui.waitForSelector('.rendered-host .sf-slide[data-slide="1"]');
  ok((await ui.getAttribute('.rendered .film-item >> nth=1', 'aria-current')) === 'true', 'filmstrip in the Rendered tab selects the slide');
  await ui.keyboard.press('ArrowUp');
  await ui.waitForSelector('.rendered-host .sf-slide[data-slide="0"]');
  ok(true, 'arrow keys move between rendered slides');
  await ui.click('.rendered .film-item >> nth=1'); await ui.waitForSelector('.rendered-host .sf-slide[data-slide="1"]');

  // the other tabs still work, and come back in sync
  await ui.click('role=tab[name="Slides"]');
  ok((await ui.locator('.slide .shp').count()) >= 1 && (await ui.getAttribute('.preview:not(.rendered) .film-item >> nth=1', 'aria-current')) === 'true', 'Slides tab still draws its structural preview, on the slide chosen in Rendered');
  ok(await ui.locator('.rendered-host').isHidden(), 'Rendered tab is hidden while Slides is open');
  await ui.click('role=tab[name="Design"]');
  ok((await ui.locator('.design .chip').count()) >= 1 && (await ui.textContent('.design')).length > 100, 'Design tab still works');
  await ui.click('role=tab[name="Rendered"]'); await ui.waitForSelector('.rendered-host .sf-slide[data-slide="1"]');
  ok(true, 'returning to Rendered keeps the selected slide');

  // responsive
  await ui.setViewportSize({ width: 820, height: 900 });
  await ui.waitForFunction(() => { const r = document.querySelector('.rendered-host .sf-slide'); const f = document.querySelector('.sf-frame'); return r && Math.abs(r.getBoundingClientRect().width - f.clientWidth) <= 1 && f.clientWidth < 800; }, null, { timeout: 5000 });
  ok(true, 'the slide re-fits when the window is resized (single --u variable)');
  await ui.setViewportSize({ width: 1280, height: 900 });
  await ui.keyboard.press('Escape');
  await ui.waitForSelector('dialog', { state: 'detached' }).catch(() => {});

  // unsupported content is reported in the tab
  await ui.click('.tpl:has-text("Unsupported deck") >> text=Preview'); await ui.waitForSelector('.film-item');
  await ui.click('role=tab[name="Rendered"]'); await ui.waitForSelector('.rendered-host .sf-slide .sf-el');
  const un = await ui.evaluate(() => ({ opaque: document.querySelectorAll('.rendered-host .sf-el[data-support="opaque"]').length, labels: [...document.querySelectorAll('.sf-opaque-label')].map((n) => n.textContent), summary: document.querySelector('.rendered .details').textContent }));
  ok(un.opaque >= 3 && un.labels.some((l) => /Chart/.test(l)) && /not drawn exactly/.test(un.summary) && /Not drawn/.test(un.summary), 'charts, SmartArt and embedded objects get a visible "not drawn" box and are listed with reasons', un);
  await ui.keyboard.press('Escape');
  await ui.close();
}


// ==== Step 8a: regression for placeholder inheritance found by the LibreOffice comparison ====
section('placeholder inheritance: layout -> master is matched by type, not idx (PowerPoint "Two Content" layout)');
{
  const real = realFiles.find((f) => /ppt-default-english/.test(f));
  if (real) {
    const r = await inPage(real, async ({ b64 }) => {
      const { parsed, scene } = await __load(b64); const rs = await import('/js/model/resolve.js');
      const s = await scene.buildScene(parsed, 2); await rs.resolveScene(s);
      return s.elements.filter((e) => e.text).map((e) => ({ ph: e.ph && `${e.ph.type}/${e.ph.idx}`, anchor: e.text.body.anchor, autofit: e.text.body.autofit, y: Math.round(e.world.cy - e.world.h / 2) }));
    });
    const bodies = r.filter((e) => /^body/.test(e.ph));
    ok(bodies.length === 2 && bodies.every((b) => b.anchor === 'top') && bodies[0].y === bodies[1].y && bodies[0].autofit === bodies[1].autofit, 'both content placeholders are top-anchored at the same height (idx 2 no longer picks up the master date placeholder)', r);
    ok(r.find((e) => e.ph === 'title/null').anchor === 'middle', 'the title keeps its middle anchor from the master');
  }
}


// ==== Step 8b: visual fidelity against LibreOffice, an independent PPTX renderer ====
import { execFileSync, spawnSync } from 'node:child_process';
section('fidelity: browser render vs LibreOffice render, slide by slide (coarse 96x54 grid, 960 px)');
{
  const have = spawnSync('sh', ['-c', 'command -v soffice && command -v pdftoppm'], { encoding: 'utf8' }).status === 0;
  if (!have) console.log('SKIPPED: soffice / pdftoppm not installed (visual fidelity vs LibreOffice)');
  else {
    const out = path.join(SHOTS, 'fidelity.json');
    const decks = renderDecks;
    const run = spawnSync('python3', ['-I', path.join(ROOT, 'tests/fidelity.py'), out, SHOTS, ...decks], { encoding: 'utf8', timeout: 600000 });
    ok(run.status === 0, 'LibreOffice rendered every deck', run.stderr?.slice(0, 300));
    const rows = run.status === 0 ? JSON.parse(readFileSync(out, 'utf8')) : [];
    // Documented waivers: places where the two renderers legitimately differ, and by how much we still require them to agree.
    const WAIVE = {
      'p31-unsupported': { mean: 9, tiles: 0.08, why: 'LibreOffice draws the chart and the orphan placeholder; we show labelled "not drawn" boxes by design' },
      'p31-media': { mean: 7, tiles: 0.05, why: 'EMF / linked / duotone pictures are placeholders by design' },
    };
    const worst = [];
    for (const r of rows) {
      const w = WAIVE[r.deck] || { mean: 6, tiles: 0.05 };
      if (r.error) { ok(false, `${r.deck} slide ${r.slide}: ${r.error}`); continue; }
      const good = r.mean <= w.mean && r.tiles_over_40 <= w.tiles && r.size[0] === r.lo_size[0] && Math.abs(r.size[1] - r.lo_size[1]) <= 2;
      worst.push(r);
      ok(good, `${r.deck} slide ${r.slide}: matches LibreOffice (mean error ${r.mean}/255, ${(r.tiles_over_40 * 100).toFixed(1)}% of tiles differ strongly)${WAIVE[r.deck] ? ' [waiver: ' + WAIVE[r.deck].why + ']' : ''}`, r);
    }
    const avg = rows.reduce((a, r) => a + (r.mean || 0), 0) / Math.max(rows.length, 1);
    ok(rows.length >= 30 && avg < 3, `average error over ${rows.length} slides is ${avg.toFixed(2)}/255`, { n: rows.length, avg });
  }
}

// ==== Step 8c: round trip ====
import { writeFileSync } from 'node:fs';
section('round trip: scene + render, then re-serialise the live XML into a new .pptx');
{
  const pairs = [];
  for (const f of renderDecks) {
    const out = await page.evaluate(async ({ b64 }) => {
      const { parsed, scene } = await globalThis.__load(b64);
      const rn = await import('/js/render/render.js'); const host = document.getElementById('sf-host');
      const docs = new Map();
      for (let i = 0; i < parsed.slides.length; i++) {
        const sc = await scene.buildScene(parsed, i);
        await rn.renderSlide(sc, host, { width: 800 }); // everything the Rendered tab does, including resolve and text fitting
        docs.set(sc.part, sc.doc);
        for (const l of sc.layers) { const d = l.elements[0]?.el.ownerDocument; if (d) docs.set(l.part, d); }
        if (sc.layout?.doc) docs.set(sc.layoutPart, sc.layout.doc);
        if (sc.master?.doc) docs.set(sc.masterPart, sc.master.doc);
      }
      const decl = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n';
      for (const [part, doc] of docs) { if (!part || !doc) continue; let x = new XMLSerializer().serializeToString(doc); if (!/^<\?xml/.test(x)) x = decl + x; parsed.zip.file(part, x, { createFolders: false }); }
      return { b64: await parsed.zip.generateAsync({ type: 'base64', compression: 'DEFLATE' }), parts: [...docs.keys()].length };
    }, { b64: b64(f) });
    const rt = path.join(SHOTS, path.basename(f, '.pptx') + '.roundtrip.pptx');
    writeFileSync(rt, Buffer.from(out.b64, 'base64'));
    pairs.push(f, rt);
  }
  const have = spawnSync('python3', ['-c', 'import lxml, pptx'], { encoding: 'utf8' }).status === 0;
  if (!have) console.log('SKIPPED: python lxml / python-pptx not installed (round-trip comparison)');
  else {
    const out = path.join(SHOTS, 'roundtrip.json');
    const run = spawnSync('python3', ['-I', path.join(ROOT, 'tests/roundtrip.py'), out, ...pairs], { encoding: 'utf8', timeout: 900000 });
    ok(run.status === 0, 'round-trip checker ran', run.stderr?.slice(0, 300));
    const rows = run.status === 0 ? JSON.parse(readFileSync(out, 'utf8')) : [];
    for (const r of rows) {
      ok(r.same_parts && r.binary_changed.length === 0, `${r.deck}: same package parts; media, charts, SmartArt, OLE and every other untouched part is byte-identical (${r.identical_bytes} parts)`, r.errors);
      ok(r.reserialised_xml === r.reserialised_equal, `${r.deck}: ${r.reserialised_equal}/${r.reserialised_xml} re-serialised XML parts are canonically identical to the original`, r.errors);
      ok(r.pptx_same === true, `${r.deck}: python-pptx reads the same shapes (ids, names, positions, rotation) and text from the round-tripped file`, r.errors);
      if (r.lo_error) ok(r.lo_error.every((e) => e <= 0.5) && r.lo_slides[0] === r.lo_slides[1], `${r.deck}: LibreOffice renders the round-tripped file identically (max pixel error ${Math.max(...r.lo_error)})`, r.errors);
    }
    ok(rows.length === renderDecks.length, `${rows.length} decks round-tripped`);
  }
}

if (!KEEP_SHOTS) rmSync(SHOTS, { recursive: true, force: true });

// ==== END (add new sections above this line) ====
section('page health');
ok(consoleErrors.length === 0, 'no console errors / CSP violations during the run', consoleErrors.slice(0, 5));
await browser.close(); server.close();
console.log(`\n${pass} passed, ${fail} failed`);
process.exit(fail ? 1 : 0);
