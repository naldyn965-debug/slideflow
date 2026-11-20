// Style resolution (Phase 3.1): turns the XML of each scene element into plain, renderer-ready values: fill, outline, effects, text (paragraphs and runs
// with every inherited property applied) and tables. Read-only: it only reads the DOM and writes `e.style`, `e.text`, `e.table` on the scene element.
// Text style resolution reuses the Phase 2 helpers exported from design.js (runStyle, alignOf, phShape...), so a run here has the same size, font and
// colour the design analysis reports. Anything simplified is recorded with note(): the element becomes `approximated`, with the reason.
import { kid, kids, desc, readXml, readRels } from '../pptx/parser.js';
import { runStyle, alignOf, RTL } from '../pptx/design.js';
import { note, relsOf, deckContext, phSources } from './scene.js';

const NS_A = 'http://schemas.openxmlformats.org/drawingml/2006/main';
const FILLS = new Set(['solidFill', 'gradFill', 'blipFill', 'pattFill', 'noFill', 'grpFill']);

// ---------- colour engine (sRGB hex <-> HSL / linear RGB, OOXML colour transforms applied in the order written) ----------
const PRST = { black: '000000', white: 'FFFFFF', red: 'FF0000', lime: '00FF00', green: '008000', blue: '0000FF', yellow: 'FFFF00', cyan: '00FFFF', aqua: '00FFFF', magenta: 'FF00FF', fuchsia: 'FF00FF', gray: '808080', grey: '808080', silver: 'C0C0C0', maroon: '800000', olive: '808000', navy: '000080', purple: '800080', teal: '008080', orange: 'FFA500', dkGray: 'A9A9A9', dkGrey: 'A9A9A9', ltGray: 'D3D3D3', ltGrey: 'D3D3D3', dimGray: '696969', dimGrey: '696969', darkBlue: '00008B', darkRed: '8B0000', darkGreen: '006400' };
const hexRgb = (h) => [0, 2, 4].map((i) => parseInt(h.slice(i, i + 2), 16));
const rgbHex = (c) => c.map((v) => Math.round(Math.min(255, Math.max(0, v))).toString(16).padStart(2, '0')).join('').toUpperCase();
const lin = (v) => { v /= 255; return v <= 0.04045 ? v / 12.92 : ((v + 0.055) / 1.055) ** 2.4; };
const unlin = (v) => 255 * (v <= 0.0031308 ? v * 12.92 : 1.055 * v ** (1 / 2.4) - 0.055);
function toHsl(c) {
  const [r, g, b] = c.map((v) => v / 255); const mx = Math.max(r, g, b); const mn = Math.min(r, g, b); const d = mx - mn; const l = (mx + mn) / 2;
  let h = 0; let s = 0;
  if (d) { s = l > 0.5 ? d / (2 - mx - mn) : d / (mx + mn); h = (mx === r ? (g - b) / d + (g < b ? 6 : 0) : mx === g ? (b - r) / d + 2 : (r - g) / d + 4) / 6; }
  return [h, s, l];
}
function fromHsl([h, s, l]) {
  h = ((h % 1) + 1) % 1; s = Math.min(1, Math.max(0, s)); l = Math.min(1, Math.max(0, l));
  if (!s) return [l * 255, l * 255, l * 255];
  const q = l < 0.5 ? l * (1 + s) : l + s - l * s; const p = 2 * l - q;
  const f = (t) => { t = (t + 1) % 1; return t < 1 / 6 ? p + (q - p) * 6 * t : t < 1 / 2 ? q : t < 2 / 3 ? p + (q - p) * (2 / 3 - t) * 6 : p; };
  return [f(h + 1 / 3) * 255, f(h) * 255, f(h - 1 / 3) * 255];
}
const pct = (el) => Number(el.getAttribute('val')) / 100000;

/** One colour element (srgbClr, schemeClr, sysClr, prstClr...) -> { hex, alpha, approx? }. `ph` is the placeholder colour (phClr) a style reference supplies. */
export function colorEl(c, cx, ph) {
  const n = c.localName; const approx = [];
  let rgb = null; let alpha = 1;
  if (n === 'srgbClr') rgb = /^[0-9a-f]{6}$/i.test(c.getAttribute('val') || '') ? hexRgb(c.getAttribute('val')) : null;
  else if (n === 'sysClr') { const v = c.getAttribute('lastClr') || ({ window: 'FFFFFF', windowText: '000000' }[c.getAttribute('val')]); rgb = v ? hexRgb(v) : null; }
  else if (n === 'prstClr') { const v = PRST[c.getAttribute('val')]; if (v) rgb = hexRgb(v); else { rgb = [0, 0, 0]; approx.push(`preset colour ${c.getAttribute('val')}`); } }
  else if (n === 'hslClr') rgb = fromHsl([Number(c.getAttribute('hue')) / 21600000, Number(c.getAttribute('sat')) / 100000, Number(c.getAttribute('lum')) / 100000]);
  else if (n === 'scrgbClr') rgb = ['r', 'g', 'b'].map((k) => unlin(Number(c.getAttribute(k)) / 100000));
  else if (n === 'schemeClr') {
    const v = c.getAttribute('val');
    if (v === 'phClr') { if (!ph) return null; rgb = hexRgb(ph.hex); alpha = ph.alpha ?? 1; } else { const slot = cx.clrMap[v] || v; const hex = cx.theme.colors[slot]; rgb = hex ? hexRgb(hex) : null; }
  }
  if (!rgb) return null;
  for (const m of c.children) {
    const k = m.localName; const v = pct(m);
    if (k === 'alpha') alpha = Math.min(1, Math.max(0, v));
    else if (k === 'alphaMod') alpha = Math.min(1, Math.max(0, alpha * v));
    else if (k === 'alphaOff') alpha = Math.min(1, Math.max(0, alpha + v));
    else if (k === 'lumMod') { const h = toHsl(rgb); h[2] *= v; rgb = fromHsl(h); }
    else if (k === 'lumOff') { const h = toHsl(rgb); h[2] += v; rgb = fromHsl(h); }
    else if (k === 'lum') { const h = toHsl(rgb); h[2] = v; rgb = fromHsl(h); }
    else if (k === 'satMod') { const h = toHsl(rgb); h[1] *= v; rgb = fromHsl(h); }
    else if (k === 'satOff') { const h = toHsl(rgb); h[1] += v; rgb = fromHsl(h); }
    else if (k === 'sat') { const h = toHsl(rgb); h[1] = v; rgb = fromHsl(h); }
    else if (k === 'hueMod') { const h = toHsl(rgb); h[0] *= v; rgb = fromHsl(h); }
    else if (k === 'hueOff') { const h = toHsl(rgb); h[0] += Number(m.getAttribute('val')) / 21600000; rgb = fromHsl(h); }
    else if (k === 'hue') { const h = toHsl(rgb); h[0] = Number(m.getAttribute('val')) / 21600000; rgb = fromHsl(h); }
    else if (k === 'shade') rgb = rgb.map((x) => unlin(lin(x) * v)); // linear-light, as LibreOffice and PowerPoint agree on
    else if (k === 'tint') rgb = rgb.map((x) => unlin(1 - (1 - lin(x)) * v));
    else approx.push(k); // comp, inv, gray, gamma, red/green/blue modifiers: not applied
  }
  return { hex: rgbHex(rgb), alpha, approx: approx.length ? approx : undefined };
}
const firstColor = (holder, cx, ph) => { const c = holder && [...holder.children].find((e) => ['srgbClr', 'schemeClr', 'sysClr', 'prstClr', 'hslClr', 'scrgbClr'].includes(e.localName)); return c ? colorEl(c, cx, ph) : null; };

// ---------- theme format scheme (fill / line / effect / background styles referenced by idx) ----------
async function themeFormats(scene) {
  const ctx = await deckContext(scene.parsed);
  const part = scene.cx.theme.part;
  if (!part) return { fills: [], lines: [], effects: [], bgFills: [] };
  ctx.fmt ||= new Map();
  if (!ctx.fmt.has(part)) {
    const doc = await readXml(ctx.zip, part);
    const fs = doc && desc(doc, 'fmtScheme')[0];
    const ch = (n) => (kid(fs, n) ? [...kid(fs, n).children] : []);
    ctx.fmt.set(part, { fills: ch('fillStyleLst'), lines: ch('lnStyleLst'), effects: ch('effectStyleLst'), bgFills: ch('bgFillStyleLst') });
  }
  return ctx.fmt.get(part);
}

// ---------- fills ----------
function fillNode(node, k) {
  const { cx, ph, rels, groupFill } = k;
  const n = node.localName;
  if (n === 'noFill') return { type: 'none' };
  if (n === 'grpFill') return groupFill || { type: 'none' };
  if (n === 'solidFill') { const c = firstColor(node, cx, ph); return c ? { type: 'solid', hex: c.hex, alpha: c.alpha, approx: c.approx } : { type: 'none', unresolved: true }; }
  if (n === 'gradFill') {
    const stops = desc(node, 'gs').map((g) => { const c = firstColor(g, cx, ph); return c ? { pos: Number(g.getAttribute('pos')) / 1000, hex: c.hex, alpha: c.alpha } : null; }).filter(Boolean).sort((a, b) => a.pos - b.pos);
    const lin = kid(node, 'lin'); const path = kid(node, 'path');
    if (!stops.length) return { type: 'none', unresolved: true };
    if (path) return { type: 'gradient', kind: 'path', path: path.getAttribute('path') || 'circle', stops, rotWithShape: node.getAttribute('rotWithShape') === '1' };
    return { type: 'gradient', kind: 'linear', angle: (Number(lin?.getAttribute('ang')) || 0) / 60000, scaled: lin?.getAttribute('scaled') === '1', stops, rotWithShape: node.getAttribute('rotWithShape') === '1' };
  }
  if (n === 'blipFill') {
    const b = kid(node, 'blip'); const rid = b && (b.getAttributeNS('http://schemas.openxmlformats.org/officeDocument/2006/relationships', 'embed') || b.getAttributeNS('http://schemas.openxmlformats.org/officeDocument/2006/relationships', 'link'));
    const rel = rid && rels?.get(rid);
    const src = kid(node, 'srcRect');
    const crop = src ? { l: Number(src.getAttribute('l')) || 0, t: Number(src.getAttribute('t')) || 0, r: Number(src.getAttribute('r')) || 0, b: Number(src.getAttribute('b')) || 0 } : null;
    const blipFx = b ? [...b.children].map((c) => c.localName).filter((x) => x !== 'extLst') : [];
    const alphaMod = b && kid(b, 'alphaModFix') ? Number(kid(b, 'alphaModFix').getAttribute('amt')) / 100000 : null;
    return { type: 'image', target: rel?.target || null, external: !!rel?.external, missing: !rel, crop, tile: !!kid(node, 'tile'), fx: blipFx, alpha: alphaMod, gray: blipFx.includes('grayscl') };
  }
  if (n === 'pattFill') { const fg = firstColor(kid(node, 'fgClr'), cx, ph); const bg = firstColor(kid(node, 'bgClr'), cx, ph); return { type: 'pattern', prst: node.getAttribute('prst'), fg: fg?.hex || '000000', bg: bg?.hex || 'FFFFFF', fgAlpha: fg?.alpha, bgAlpha: bg?.alpha }; }
  return null;
}
const fillChild = (holder) => (holder ? [...holder.children].find((c) => FILLS.has(c.localName)) || null : null);

function noteFill(e, f, what) {
  if (!f) return;
  if (f.type === 'pattern') note(e, 'approximated', `${what}: pattern fill (${f.prst}) is drawn as a flat blend of its two colours`);
  else if (f.type === 'gradient' && f.kind === 'path') note(e, 'approximated', `${what}: ${f.path} gradient is drawn as a radial/linear approximation`);
  else if (f.type === 'image' && f.tile) note(e, 'approximated', `${what}: tiled picture fill is drawn stretched`);
  else if (f.type === 'image' && f.fx?.some((x) => !['alphaModFix', 'grayscl'].includes(x))) note(e, 'approximated', `${what}: picture effect (${f.fx.filter((x) => !['alphaModFix', 'grayscl'].includes(x)).join(', ')}) is not applied`);
  if (f?.approx) note(e, 'approximated', `${what}: colour transform ${f.approx.join(', ')} is not applied`);
}

// ---------- outline ----------
const DASH = { solid: [], dot: [1, 3], sysDot: [1, 1], dash: [4, 3], sysDash: [3, 1], lgDash: [8, 3], dashDot: [4, 3, 1, 3], sysDashDot: [3, 1, 1, 1], lgDashDot: [8, 3, 1, 3], lgDashDotDot: [8, 3, 1, 3, 1, 3], sysDashDotDot: [3, 1, 1, 1, 1, 1] };
function readLn(ln, k, into) {
  if (!ln) return into;
  if (ln.hasAttribute('w')) into.w = Number(ln.getAttribute('w')) || 0;
  if (ln.hasAttribute('cap')) into.cap = { rnd: 'round', sq: 'square', flat: 'butt' }[ln.getAttribute('cap')] || 'butt';
  if (ln.hasAttribute('cmpd') && ln.getAttribute('cmpd') !== 'sng') into.cmpd = ln.getAttribute('cmpd');
  const f = fillChild(ln);
  if (f) {
    const r = fillNode(f, k);
    if (r?.type === 'solid') Object.assign(into, { type: 'solid', hex: r.hex, alpha: r.alpha, approx: r.approx });
    else if (r?.type === 'gradient') Object.assign(into, { type: 'solid', hex: r.stops[0].hex, alpha: r.stops[0].alpha, gradientApprox: true });
    else if (r?.type === 'pattern') Object.assign(into, { type: 'solid', hex: r.fg, patternApprox: true });
    else into.type = 'none';
  }
  const d = kid(ln, 'prstDash'); if (d) into.dash = d.getAttribute('val') || 'solid';
  if (kid(ln, 'custDash')) into.customDash = true;
  if (kid(ln, 'round')) into.join = 'round'; else if (kid(ln, 'bevel')) into.join = 'bevel'; else if (kid(ln, 'miter')) into.join = 'miter';
  for (const [side, tag] of [['head', 'headEnd'], ['tail', 'tailEnd']]) { const h = kid(ln, tag); if (h && h.getAttribute('type') && h.getAttribute('type') !== 'none') into[side] = { type: h.getAttribute('type'), w: h.getAttribute('w') || 'med', len: h.getAttribute('len') || 'med' }; else if (h) into[side] = null; }
  return into;
}

// ---------- the shared per-element resolution ----------
function placeholderSpPr(scene, e) {
  if (!e.ph) return [];
  const { layout, master } = phSources(e.ph, scene.layout?.doc, scene.master?.doc, e.origin);
  return [layout, master].filter(Boolean).map((s) => kid(s, 'spPr')).filter(Boolean);
}

function resolveShapeStyle(scene, e, fmt, groupFills) {
  const el = e.el; const spPr = kid(el, e.tag === 'grpSp' ? 'grpSpPr' : 'spPr');
  const rels = relsOf(scene, e);
  const style = kid(el, 'style');
  const k = { cx: scene.cx, rels, groupFill: null };
  const parentGroup = e.parentId && scene.byId.get(e.parentId);
  k.groupFill = parentGroup ? groupFills.get(parentGroup.id) || null : null;
  const out = { fill: { type: 'none' }, line: { type: 'none' }, shadow: null };

  // ---- fill: own, else placeholder source, else <p:style> fillRef ----
  let src = fillChild(spPr) || placeholderSpPr(scene, e).map(fillChild).find(Boolean) || null;
  const ref = kid(style, 'fillRef');
  if (src) out.fill = fillNode(src, k) || out.fill;
  else if (ref && ref.getAttribute('idx') !== '0') {
    const idx = Number(ref.getAttribute('idx')) || 0;
    const phc = firstColor(ref, scene.cx, null);
    const node = idx >= 1001 ? fmt.bgFills[idx - 1001] : fmt.fills[idx - 1];
    if (node) out.fill = fillNode(node, { ...k, ph: phc }) || out.fill;
    else if (phc) { out.fill = { type: 'solid', hex: phc.hex, alpha: phc.alpha }; note(e, 'approximated', `theme fill style ${idx} is missing; the reference colour is used`); }
  }
  if (e.tag === 'grpSp') groupFills.set(e.id, out.fill);
  noteFill(e, out.fill, 'fill');

  // ---- outline: theme line style (lnRef) as the base, own <a:ln> (or the placeholder's) on top ----
  const ln = { type: 'none' };
  const lref = kid(style, 'lnRef');
  if (lref && lref.getAttribute('idx') !== '0') {
    const idx = Number(lref.getAttribute('idx')) || 0; const phc = firstColor(lref, scene.cx, null); const node = fmt.lines[idx - 1];
    if (node) readLn(node, { ...k, ph: phc }, ln); else if (phc) { Object.assign(ln, { type: 'solid', hex: phc.hex, alpha: phc.alpha, w: 12700 }); }
  }
  const phLn = placeholderSpPr(scene, e).map((p) => kid(p, 'ln')).find(Boolean);
  if (!kid(spPr, 'ln') && phLn) readLn(phLn, k, ln);
  readLn(kid(spPr, 'ln'), k, ln);
  if (ln.type === 'solid' && ln.w == null) ln.w = 9525;
  if (ln.type === 'solid' && ln.w === 0) ln.w = 9525; // PowerPoint draws a hairline as 0.75 pt
  out.line = ln;
  if (ln.cmpd) note(e, 'approximated', `outline: compound line (${ln.cmpd}) is drawn as one line`);
  if (ln.gradientApprox) note(e, 'approximated', 'outline: gradient line is drawn with its first colour');
  if (ln.patternApprox) note(e, 'approximated', 'outline: pattern line is drawn as a solid line');
  if (ln.customDash) note(e, 'approximated', 'outline: custom dash pattern is drawn as a standard dash');
  if (ln.approx) note(e, 'approximated', `outline: colour transform ${ln.approx.join(', ')} is not applied`);

  // ---- effects: own <a:effectLst>, else the theme effect style named by effectRef ----
  let eff = spPr && kid(spPr, 'effectLst');
  let ph = null;
  const eref = kid(style, 'effectRef');
  if (!eff && eref && eref.getAttribute('idx') !== '0') {
    const node = fmt.effects[(Number(eref.getAttribute('idx')) || 0) - 1];
    eff = node && kid(node, 'effectLst');
    ph = firstColor(eref, scene.cx, null);
    if (node && (kid(node, 'scene3d') || kid(node, 'sp3d'))) note(e, 'approximated', 'theme effect style has 3D properties (bevel / lighting) that are not drawn');
  }
  if (spPr && (kid(spPr, 'scene3d') || kid(spPr, 'sp3d'))) note(e, 'approximated', '3D properties (bevel / extrusion) are not drawn');
  for (const x of eff ? eff.children : []) {
    if (x.localName === 'outerShdw') {
      const c = firstColor(x, scene.cx, ph); const dist = Number(x.getAttribute('dist')) || 0; const dir = ((Number(x.getAttribute('dir')) || 0) / 60000 * Math.PI) / 180;
      const sx = x.getAttribute('sx') || x.getAttribute('sy') || x.getAttribute('kx') || x.getAttribute('ky');
      out.shadow = { dx: dist * Math.cos(dir), dy: dist * Math.sin(dir), blur: Number(x.getAttribute('blurRad')) || 0, hex: c?.hex || '000000', alpha: c?.alpha ?? 1 };
      if (sx) note(e, 'approximated', 'outer shadow: scale / skew of the shadow is not applied');
    } else if (x.localName === 'glow') {
      const c = firstColor(x, scene.cx, ph);
      out.glow = { rad: Number(x.getAttribute('rad')) || 0, hex: c?.hex || 'FFFFFF', alpha: c?.alpha ?? 1 };
      note(e, 'approximated', 'glow is drawn as a blurred halo, which is close to but not identical to PowerPoint\'s glow');
    } else note(e, 'approximated', `effect ${x.localName} is not drawn`);
  }
  return out;
}

// ---------- text ----------
const FOOTER = new Set(['dt', 'ftr', 'sldNum']);
const BULLET_SYMBOLS = { Wingdings: { '§': '▪', 'q': '❑', 'v': '❖', 'Ø': '➢', 'ü': '✓', 'l': '●', 'n': '■', 'u': '◆', 'þ': '☑', 'o': '○', 'x': '✗', '': '▪', '': '❖', '': '➢', '': '✓' }, Symbol: { '·': '•', '': '•' }, 'Wingdings 3': { 'u': '▶', 'p': '▲' } };
const ROMAN = [[1000, 'm'], [900, 'cm'], [500, 'd'], [400, 'cd'], [100, 'c'], [90, 'xc'], [50, 'l'], [40, 'xl'], [10, 'x'], [9, 'ix'], [5, 'v'], [4, 'iv'], [1, 'i']];
const roman = (n) => { let s = ''; for (const [v, r] of ROMAN) while (n >= v) { s += r; n -= v; } return s; };
const alpha = (n) => { let s = ''; while (n > 0) { n--; s = String.fromCharCode(97 + (n % 26)) + s; n = Math.floor(n / 26); } return s; };
function autoNumber(scheme, n) {
  const base = /^arabic/.test(scheme) ? String(n) : /^romanUc/.test(scheme) ? roman(n).toUpperCase() : /^romanLc/.test(scheme) ? roman(n) : /^alphaUc/.test(scheme) ? alpha(n).toUpperCase() : /^alphaLc/.test(scheme) ? alpha(n) : null;
  if (base == null) return null;
  return /ParenBoth$/.test(scheme) ? `(${base})` : /ParenR$/.test(scheme) ? `${base})` : /Plain$/.test(scheme) ? base : `${base}.`;
}

/** Split a run into script segments so Arabic words use the complex-script font and Latin words the Latin font (neutral characters follow the previous word). */
function splitScripts(text) {
  if (!RTL.test(text)) return [{ text, rtl: false }];
  const out = []; let cur = null;
  for (const tok of text.split(/(\s+)/)) {
    if (tok === '') continue;
    const r = RTL.test(tok);
    const neutral = !r && !/[A-Za-zÀ-ɏ]/.test(tok);
    const rtl = r || (neutral && cur ? cur.rtl : false);
    if (cur && cur.rtl === rtl) cur.text += tok; else { cur = { text: tok, rtl }; out.push(cur); }
  }
  return out;
}

function pickLayers(layers, f) { for (const l of layers) { const v = l && f(l); if (v != null) return v; } return null; }

/** Resolve one <a:txBody>-like container (p:txBody on shapes, a:txBody in table cells). `phLayers`: layout/master placeholder txBodies for inheritance. */
function resolveText(e, scene, body, opts) {
  const { ph, phBodies = [], defaultAnchor = 't', cellStyle = null } = opts;
  const env = { cx: scene.cx };
  const master = scene.master; const ctxDefault = scene.defaultTextStyle;
  const bodyPr = kid(body, 'bodyPr');
  const bodyPrs = [bodyPr, ...phBodies.map((b) => kid(b, 'bodyPr'))];
  const attr = (a) => pickLayers(bodyPrs, (b) => (b.hasAttribute(a) ? b.getAttribute(a) : null));
  const autofitNode = pickLayers(bodyPrs, (b) => ['noAutofit', 'normAutofit', 'spAutoFit'].map((t) => kid(b, t)).find(Boolean) || null);
  const fontScale = autofitNode?.localName === 'normAutofit' ? (Number(autofitNode.getAttribute('fontScale')) || 100000) / 100000 : 1;
  const lnRed = autofitNode?.localName === 'normAutofit' ? (Number(autofitNode.getAttribute('lnSpcReduction')) || 0) / 100000 : 0;
  const t = ph ? (ph.type === 'ctrTitle' ? 'title' : ph.type) : null;
  const group = !ph || FOOTER.has(ph.type) ? 'otherStyle' : t === 'title' ? 'titleStyle' : 'bodyStyle';
  const chain = [kid(body, 'lstStyle'), ...phBodies.map((b) => kid(b, 'lstStyle')), master?.styles[group], scene.defaultTextStyle].filter(Boolean);
  const fontRef = kid(kid(e.el, 'style'), 'fontRef');
  const fontCol = fontRef ? colorEl0(fontRef, scene.cx) : null;
  const shapeLayers = kid(body, 'lstStyle') ? 1 : 0;

  const bodyOut = {
    anchor: { t: 'top', ctr: 'middle', b: 'bottom', just: 'top', dist: 'top' }[attr('anchor')] || { t: 'top', ctr: 'middle', b: 'bottom' }[defaultAnchor],
    anchorCtr: attr('anchorCtr') === '1',
    insets: { l: Number(attr('lIns') ?? 91440), t: Number(attr('tIns') ?? 45720), r: Number(attr('rIns') ?? 91440), b: Number(attr('bIns') ?? 45720) },
    wrap: attr('wrap') === 'none' ? 'none' : 'square', vert: attr('vert') && attr('vert') !== 'horz' ? attr('vert') : null,
    cols: Number(attr('numCol')) || 1, colGap: Number(attr('spcCol')) || 0, rtlCol: attr('rtlCol') === '1',
    autofit: autofitNode ? ({ noAutofit: 'none', normAutofit: 'shrink', spAutoFit: 'resize' }[autofitNode.localName]) : 'none', fontScale, lnSpcReduction: lnRed,
  };
  // "shrink on overflow" without a stored scale is only an approximation when the text really overflows: the renderer checks that and calls note() itself.
  if (bodyOut.autofit === 'resize') note(e, 'approximated', 'text is set to "resize shape to fit text": the frame is drawn at the size stored in the file');
  if (bodyOut.vert) note(e, 'approximated', `vertical text (${bodyOut.vert}) is drawn with CSS writing modes`);
  if (bodyOut.cols > 1) note(e, 'approximated', 'text columns: the browser breaks lines and columns, so a column may break at a slightly different word than in PowerPoint');

  const paras = []; const counters = [];
  for (const p of kids(body, 'p')) {
    const pPr = kid(p, 'pPr');
    const lvl0 = Number(pPr?.getAttribute('lvl')) || 0; const lvl = lvl0 + 1;
    const layers = [pPr, ...chain.map((c) => kids(c, `lvl${lvl}pPr`)[0])];
    const pa = (a) => pickLayers(layers, (l) => (l.hasAttribute(a) ? l.getAttribute(a) : null));
    const rtl = pa('rtl') === '1';
    const space = (tag) => { const c = pickLayers(layers, (l) => kid(l, tag)); const pc = c && kid(c, 'spcPct'); const pt = c && kid(c, 'spcPts'); return pc ? { pct: Number(pc.getAttribute('val')) / 100000 } : pt ? { pts: Number(pt.getAttribute('val')) / 100 } : null; };
    const para = { lvl: lvl0, algn: alignOf(pPr, lvl, chain), rtl, marL: Number(pa('marL')) || 0, indent: Number(pa('indent')) || 0, lnSpc: space('lnSpc'), spcBef: space('spcBef'), spcAft: space('spcAft'), bullet: null, runs: [], endSz: null };
    if (para.lnSpc?.pct != null && lnRed) para.lnSpc = { pct: Math.max(0, para.lnSpc.pct - lnRed) };
    else if (!para.lnSpc && lnRed) para.lnSpc = { pct: 1 - lnRed };
    // bullets: the first layer that says anything about them decides
    const buLayer = layers.find((l) => l && ['buNone', 'buChar', 'buAutoNum', 'buBlip'].some((t) => kid(l, t)));
    const bu = (tag) => pickLayers(layers, (l) => kid(l, tag));
    if (buLayer && !kid(buLayer, 'buNone')) {
      const colEl = bu('buClr'); const col = colEl ? firstColor(colEl, scene.cx, null) : null;
      const szP = bu('buSzPct'); const szPts = bu('buSzPts'); const font = bu('buFont');
      const common = { color: col?.hex || null, sizePct: szP ? Number(szP.getAttribute('val')) / 100000 : null, sizePts: szPts ? Number(szPts.getAttribute('val')) / 100 : null, font: font?.getAttribute('typeface') || null };
      const ch = kid(buLayer, 'buChar'); const auto = kid(buLayer, 'buAutoNum');
      if (ch) {
        let c = ch.getAttribute('char') || '•'; const sym = BULLET_SYMBOLS[common.font]; let approx = false;
        if (sym) { if (sym[c]) { c = sym[c]; common.font = null; } else if (/^(Wingdings|Symbol)/.test(common.font)) { c = '•'; approx = true; } }
        para.bullet = { type: 'char', char: c, ...common };
        if (approx) note(e, 'approximated', `bullet from the ${common.font} symbol font has no Unicode equivalent here; a standard bullet is drawn`);
      } else if (auto) {
        const scheme = auto.getAttribute('type') || 'arabicPeriod'; const start = Number(auto.getAttribute('startAt')) || 1;
        counters.length = lvl0 + 1; counters[lvl0] = (counters[lvl0] || 0) + 1;
        if (!(counters[lvl0] > 1)) counters[lvl0] = start; else if (counters[lvl0] === 1) counters[lvl0] = start;
        const label = autoNumber(scheme, counters[lvl0]);
        para.bullet = { type: 'auto', scheme, label: label ?? autoNumber('arabicPeriod', counters[lvl0]), ...common };
        if (label == null) note(e, 'approximated', `numbering scheme ${scheme} is drawn as 1. 2. 3.`);
      } else { para.bullet = { type: 'char', char: '•', ...common }; note(e, 'approximated', 'picture bullet is drawn as a standard bullet'); }
    }
    if (!para.bullet) counters.length = Math.min(counters.length, lvl0); // a non-numbered paragraph at this level restarts deeper counters
    if (!para.bullet || para.bullet.type !== 'auto') { if (counters.length > lvl0) counters.length = lvl0; }

    for (const n of p.children) {
      const ln = n.localName;
      if (ln === 'br') { para.runs.push({ br: true }); continue; }
      if (ln === 'endParaRPr') { para.endSz = Number(n.getAttribute('sz')) ? Number(n.getAttribute('sz')) / 100 : null; continue; }
      if (ln !== 'r' && ln !== 'fld') continue;
      const text = kid(n, 't')?.textContent ?? '';
      if (!text) continue;
      const rPr = kid(n, 'rPr');
      const rl = [rPr, ...chain.map((c) => kid(kids(c, `lvl${lvl}pPr`)[0], 'defRPr'))];
      const ra = (a) => pickLayers(rl, (l) => (l.hasAttribute(a) ? l.getAttribute(a) : null));
      const link = !!(rPr && kid(rPr, 'hlinkClick'));
      const hl = rPr && kid(rPr, 'highlight') ? firstColor(kid(rPr, 'highlight'), scene.cx, null) : null;
      for (const seg of splitScripts(text)) {
        const st = runStyle(rPr, lvl, chain, env, seg.rtl, fontScale, fontCol, shapeLayers);
        const u = ra('u'); const strike = ra('strike'); const base = Number(ra('baseline')) || 0; const cap = ra('cap'); const spc = Number(ra('spc')) || 0;
        let hex = st.hex; let alphaV = st.alpha;
        if (cellStyle?.color && !(rPr && kid(rPr, 'solidFill')) && !(rl[1] && kid(rl[1], 'solidFill'))) { // table-style text colour wins over the deck default, never over the cell's own colour
          hex = cellStyle.color.hex; alphaV = cellStyle.color.alpha; }
        if (link) { const h = scene.cx.theme.colors.hlink; if (h) hex = h; }
        const run = { text: seg.text, rtl: seg.rtl, sz: st.sz, b: cellStyle?.b && !rPr?.hasAttribute('b') && !st.b ? true : st.b, i: st.i, font: st.font, hex, alpha: alphaV, u: !!u && u !== 'none' || link, strike: !!strike && strike !== 'noStrike', baseline: base, caps: cap === 'all' ? 'all' : cap === 'small' ? 'small' : null, spc: spc / 100, highlight: hl?.hex || null, link };
        if (st.approx) note(e, 'approximated', `text colour transform ${st.approx.join(', ')} is not applied`);
        para.runs.push(run);
      }
      if (rPr && (kid(rPr, 'effectLst') || kid(rPr, 'gradFill') || kid(rPr, 'ln') || kid(rPr, 'blipFill') || kid(rPr, 'pattFill'))) note(e, 'approximated', 'text effect (outline, gradient or shadow on the letters) is drawn as plain text');
    }
    paras.push(para);
  }
  return { body: bodyOut, paras };
}
// colour of a fontRef (shape-style text colour)
const colorEl0 = (holder, cx) => firstColor(holder, cx, null);

// ---------- tables ----------
const BUILTIN_STYLE_XML = `<a:tblStyle xmlns:a="${NS_A}" styleId="{5C22544A-7EE6-4342-B048-85BDC9FD1C3A}" styleName="Medium Style 2 - Accent 1">
<a:wholeTbl><a:tcTxStyle><a:schemeClr val="dk1"/></a:tcTxStyle><a:tcStyle><a:tcBdr>${['left', 'right', 'top', 'bottom', 'insideH', 'insideV'].map((s) => `<a:${s}><a:ln w="12700" cmpd="sng"><a:solidFill><a:schemeClr val="lt1"/></a:solidFill></a:ln></a:${s}>`).join('')}</a:tcBdr><a:fill><a:solidFill><a:schemeClr val="accent1"><a:tint val="20000"/></a:schemeClr></a:solidFill></a:fill></a:tcStyle></a:wholeTbl>
<a:band1H><a:tcStyle><a:tcBdr/><a:fill><a:solidFill><a:schemeClr val="accent1"><a:tint val="40000"/></a:schemeClr></a:solidFill></a:fill></a:tcStyle></a:band1H>
<a:band1V><a:tcStyle><a:tcBdr/><a:fill><a:solidFill><a:schemeClr val="accent1"><a:tint val="40000"/></a:schemeClr></a:solidFill></a:fill></a:tcStyle></a:band1V>
<a:lastCol><a:tcTxStyle b="on"><a:schemeClr val="lt1"/></a:tcTxStyle><a:tcStyle><a:tcBdr/><a:fill><a:solidFill><a:schemeClr val="accent1"/></a:solidFill></a:fill></a:tcStyle></a:lastCol>
<a:firstCol><a:tcTxStyle b="on"><a:schemeClr val="lt1"/></a:tcTxStyle><a:tcStyle><a:tcBdr/><a:fill><a:solidFill><a:schemeClr val="accent1"/></a:solidFill></a:fill></a:tcStyle></a:firstCol>
<a:lastRow><a:tcTxStyle b="on"><a:schemeClr val="lt1"/></a:tcTxStyle><a:tcStyle><a:tcBdr><a:top><a:ln w="38100" cmpd="sng"><a:solidFill><a:schemeClr val="lt1"/></a:solidFill></a:ln></a:top></a:tcBdr><a:fill><a:solidFill><a:schemeClr val="accent1"/></a:solidFill></a:fill></a:tcStyle></a:lastRow>
<a:firstRow><a:tcTxStyle b="on"><a:schemeClr val="lt1"/></a:tcTxStyle><a:tcStyle><a:tcBdr><a:bottom><a:ln w="38100" cmpd="sng"><a:solidFill><a:schemeClr val="lt1"/></a:solidFill></a:ln></a:bottom></a:tcBdr><a:fill><a:solidFill><a:schemeClr val="accent1"/></a:solidFill></a:fill></a:tcStyle></a:firstRow></a:tblStyle>`;
let builtin = null;
const builtinStyle = () => (builtin ||= new DOMParser().parseFromString(BUILTIN_STYLE_XML, 'application/xml').documentElement);

async function tableStyles(scene) {
  const ctx = await deckContext(scene.parsed);
  if (!ctx.tblStyles) {
    ctx.tblStyles = new Map();
    const presRels = await readRels(ctx.zip, 'ppt/presentation.xml');
    const r = [...presRels.values()].find((x) => x.kind === 'tableStyles');
    const doc = r && (await readXml(ctx.zip, r.target));
    for (const s of doc ? desc(doc, 'tblStyle') : []) ctx.tblStyles.set(s.getAttribute('styleId'), s);
  }
  return ctx.tblStyles;
}

function resolveTable(scene, e, fmt, styles) {
  const tbl = desc(e.el, 'tbl')[0];
  if (!tbl) { note(e, 'opaque', 'table without a table body'); return; }
  const tblPr = kid(tbl, 'tblPr');
  const flag = (a) => tblPr?.getAttribute(a) === '1';
  const styleId = kid(tblPr, 'tableStyleId')?.textContent?.trim() || null;
  let style = styleId ? styles.get(styleId) : null;
  if (!style && styleId === '{5C22544A-7EE6-4342-B048-85BDC9FD1C3A}') { style = builtinStyle(); note(e, 'approximated', 'table style "Medium Style 2 - Accent 1" is built into PowerPoint and not stored in the file: drawn from its published definition'); }
  else if (!style && styleId) note(e, 'approximated', `table style ${styleId} is built into PowerPoint and not stored in the file: its fills and borders are not applied`);
  const cols = kids(kid(tbl, 'tblGrid'), 'gridCol').map((g) => Number(g.getAttribute('w')) || 0);
  const trs = kids(tbl, 'tr');
  const rels = relsOf(scene, e);
  const k = { cx: scene.cx, rels, groupFill: null };

  const part = (name) => style && kid(style, name);
  const styleCell = (r, c, nr, nc) => {
    const parts = ['wholeTbl'];
    const body = r - (flag('firstRow') ? 1 : 0);
    if (flag('bandRow') && body >= 0 && !(flag('lastRow') && r === nr - 1)) parts.push(body % 2 === 0 ? 'band1H' : 'band2H');
    const bc = c - (flag('firstCol') ? 1 : 0);
    if (flag('bandCol') && bc >= 0) parts.push(bc % 2 === 0 ? 'band1V' : 'band2V');
    if (flag('firstCol') && c === 0) parts.push('firstCol');
    if (flag('lastCol') && c === nc - 1) parts.push('lastCol');
    if (flag('firstRow') && r === 0) parts.push('firstRow');
    if (flag('lastRow') && r === nr - 1) parts.push('lastRow');
    const out = { fill: null, b: false, color: null, bdr: {} };
    for (const nme of parts) {
      const p = part(nme); if (!p) continue;
      const tx = kid(p, 'tcTxStyle');
      if (tx) { if (tx.getAttribute('b') === 'on') out.b = true; else if (tx.getAttribute('b') === 'off') out.b = false; const col = firstColor(tx, scene.cx, null); if (col) out.color = col; }
      const ts = kid(p, 'tcStyle'); if (!ts) continue;
      const f = kid(kid(ts, 'fill'), 'solidFill'); if (f) { const fr = fillNode(f, k); if (fr) out.fill = fr; }
      for (const side of ['left', 'right', 'top', 'bottom', 'insideH', 'insideV']) { const b = kid(kid(ts, 'tcBdr'), side); const l = b && kid(b, 'ln'); if (l) out.bdr[side] = readLn(l, k, {}); }
    }
    return out;
  };

  const nr = trs.length; const nc = cols.length;
  const rows = trs.map((tr, r) => ({ h: Number(tr.getAttribute('h')) || 0, cells: kids(tr, 'tc').map((tc, c) => {
    const tcPr = kid(tc, 'tcPr'); const sc = styleCell(r, c, nr, nc);
    const own = fillChild(tcPr); let fill = sc.fill;
    if (own) fill = fillNode(own, k);
    const bdr = { ...sc.bdr };
    for (const [attr, side] of [['lnL', 'left'], ['lnR', 'right'], ['lnT', 'top'], ['lnB', 'bottom']]) { const l = kid(tcPr, attr); if (l) bdr[side] = readLn(l, k, {}); }
    const body = kid(tc, 'txBody');
    const text = body ? resolveText(e, scene, body, { ph: null, defaultAnchor: 't', cellStyle: { b: sc.b, color: sc.color } }) : { body: null, paras: [] };
    const pr = (a, d) => (tcPr?.hasAttribute(a) ? Number(tcPr.getAttribute(a)) : d);
    return { rowSpan: Number(tc.getAttribute('rowSpan')) || 1, gridSpan: Number(tc.getAttribute('gridSpan')) || 1, hMerge: tc.getAttribute('hMerge') === '1', vMerge: tc.getAttribute('vMerge') === '1',
      fill, bdr, text: text.paras, anchor: { t: 'top', ctr: 'middle', b: 'bottom' }[tcPr?.getAttribute('anchor')] || 'top', mar: { l: pr('marL', 91440), r: pr('marR', 91440), t: pr('marT', 45720), b: pr('marB', 45720) }, vert: tcPr?.getAttribute('vert') || null };
  }) }));
  if (rows.some((r) => r.cells.some((c) => c.vert && c.vert !== 'horz'))) note(e, 'approximated', 'vertical text in table cells is drawn horizontally');
  e.table = { cols, rows, styleId, rtl: tblPr?.getAttribute('rtl') === '1', flags: { firstRow: flag('firstRow'), lastRow: flag('lastRow'), firstCol: flag('firstCol'), lastCol: flag('lastCol'), bandRow: flag('bandRow'), bandCol: flag('bandCol') } };
}

// ---------- background ----------
function resolveBackground(scene, fmt) {
  const bg = scene.background;
  const cx = scene.cx;
  const dflt = cx.theme.colors[cx.clrMap.bg1];
  if (!bg.el) return { type: 'solid', hex: dflt || 'FFFFFF', alpha: 1, source: 'default' };
  const k = { cx, rels: bg.rels, groupFill: null };
  const pr = kid(bg.el, 'bgPr');
  let f = null;
  if (pr) { const n = fillChild(pr); f = n && fillNode(n, k); }
  else {
    const ref = kid(bg.el, 'bgRef'); const idx = Number(ref?.getAttribute('idx')) || 0; const phc = firstColor(ref, cx, null);
    const node = idx >= 1001 ? fmt.bgFills[idx - 1001] : idx >= 1 ? fmt.fills[idx - 1] : null;
    f = node ? fillNode(node, { ...k, ph: phc }) : phc ? { type: 'solid', hex: phc.hex, alpha: phc.alpha } : null;
  }
  return { ...(f || { type: 'solid', hex: dflt || 'FFFFFF', alpha: 1 }), source: bg.source };
}

/** Resolve every element of the scene (slide, master and layout layers) and the background. Idempotent. */
export async function resolveScene(scene) {
  if (scene.resolved) return scene;
  const ctx = await deckContext(scene.parsed);
  scene.defaultTextStyle = ctx.env.defaultTextStyle;
  const fmt = await themeFormats(scene);
  const styles = await tableStyles(scene);
  const groupFills = new Map();
  scene.backgroundFill = resolveBackground(scene, fmt);
  noteBg(scene);
  for (const layer of scene.layers) {
    for (const e of layer.elements) {
      if (e.kind === 'unknown' || e.kind === 'alt' || e.kind === 'content') continue;
      if (e.kind === 'chart' || e.kind === 'diagram' || e.kind === 'ole' || e.kind === 'frame') continue;
      if (e.tag === 'graphicFrame') { resolveTable(scene, e, fmt, styles); continue; }
      e.style = resolveShapeStyle(scene, e, fmt, groupFills);
      if (e.tag === 'pic') {
        const blipFill = kid(e.el, 'blipFill');
        const f = blipFill ? fillNode(blipFill, { cx: scene.cx, rels: relsOf(scene, e), groupFill: null }) : null;
        e.style.fill = f && f.type === 'image' ? f : { type: 'none' };
        e.image = { ...(e.image || {}), crop: f?.crop || null, fx: f?.fx || [], alpha: f?.alpha ?? null, gray: !!f?.gray };
        noteFill(e, e.style.fill, 'picture');
        if (f && f.type === 'image') {
          if (f.external) note(e, 'approximated', 'linked picture (stored outside the file): drawn as a placeholder, nothing is fetched from the network');
          else if (f.missing) note(e, 'approximated', 'the picture part is missing from the file: drawn as a placeholder');
          else if (/\.(emf|wmf|tiff?|wdp|hdp|jxr)$/i.test(f.target || '')) note(e, 'approximated', `picture format ${f.target.split('.').pop().toUpperCase()} cannot be shown by the browser: drawn as a placeholder, the file is untouched`);
        } else if (blipFill) note(e, 'approximated', 'picture could not be read: drawn as a placeholder');
      }
      if (e.tag === 'sp') {
        const body = kid(e.el, 'txBody');
        if (body) {
          const srcs = e.ph ? phSources(e.ph, scene.layout?.doc, scene.master?.doc, e.origin) : { layout: null, master: null };
          const phBodies = [srcs.layout, srcs.master].filter(Boolean).map((s) => kid(s, 'txBody')).filter(Boolean);
          e.text = resolveText(e, scene, body, { ph: e.ph, phBodies, defaultAnchor: 't' });
        }
      }
    }
  }
  scene.resolved = true;
  return scene;
}
function noteBg(scene) { /* background effects are reported by the renderer through scene.backgroundFill */ }
