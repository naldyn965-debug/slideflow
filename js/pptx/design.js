// Template intelligence (Phase 2): theme, text styles, backgrounds, layout/spacing statistics, slide classification
// and the JSON "Design System". Read-only and rule-based: every number comes from the file's XML, every class from
// stated rules with a confidence. Anything that cannot be resolved is reported in `limitations`, never guessed.
import { kid, kids, desc, readXml, readRels, xfrmOf, intoGroup, placeholderOf, normType } from './parser.js';

// v2 (Phase 2.1): colour-map overrides, per-shape backdrops for contrast, bleed-aware margins, background-aware classification.
// v3 (Phase 2.2): rotation-aware geometry, decorative bleed vs content overflow, text-density/picture-role/structure-aware slide types.
export const DESIGN_VERSION = 3;
const NS_R = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
const FOOTER = new Set(['dt', 'ftr', 'sldNum']);
const DEFAULT_CLRMAP = { bg1: 'lt1', tx1: 'dk1', bg2: 'lt2', tx2: 'dk2' };
const RTL = /[\u0590-\u08FF\uFB1D-\uFDFF\uFE70-\uFEFF]/;
const EMPTY_THEME = { name: null, part: null, colors: {}, fonts: {}, missing: true };
const ALIGN = { l: 'left', ctr: 'center', r: 'right', just: 'justify', dist: 'justify' };
export const CLASSES = ['title', 'section', 'text', 'image', 'comparison', 'chart-table', 'closing', 'other'];

// ---------- small utils ----------
const median = (a) => { if (!a.length) return null; const s = [...a].sort((x, y) => x - y); const m = s.length >> 1; return s.length % 2 ? s[m] : (s[m - 1] + s[m]) / 2; };
const pctile = (a, p) => { if (!a.length) return null; const s = [...a].sort((x, y) => x - y); return s[Math.min(s.length - 1, Math.floor(p * (s.length - 1) + 0.5))]; };
const r1 = (n) => (n == null ? null : Math.round(n * 10) / 10);
const r2 = (n) => (n == null ? null : Math.round(n * 100) / 100);
const topKey = (map) => [...map.entries()].sort((a, b) => b[1] - a[1])[0]?.[0] ?? null;
const bump = (map, k, v = 1) => map.set(k, (map.get(k) || 0) + v);
const tick = () => new Promise((r) => setTimeout(r));

// ---------- colour ----------
const toRgb = (h) => [0, 2, 4].map((i) => parseInt(h.slice(i, i + 2), 16));
const toHex = (rgb) => rgb.map((v) => Math.round(Math.min(255, Math.max(0, v))).toString(16).padStart(2, '0')).join('').toUpperCase();
function lumTransform(hex, mod, off) {
  const [r, g, b] = toRgb(hex).map((v) => v / 255);
  const max = Math.max(r, g, b); const min = Math.min(r, g, b); const d = max - min;
  let l = (max + min) / 2; let s = 0; let h = 0;
  if (d) {
    s = l > 0.5 ? d / (2 - max - min) : d / (max + min);
    h = (max === r ? (g - b) / d + (g < b ? 6 : 0) : max === g ? (b - r) / d + 2 : (r - g) / d + 4) / 6;
  }
  l = Math.min(1, Math.max(0, l * mod + off));
  if (!s) return toHex([l * 255, l * 255, l * 255]);
  const q = l < 0.5 ? l * (1 + s) : l + s - l * s; const p = 2 * l - q;
  const f = (t) => { t = (t + 1) % 1; return t < 1 / 6 ? p + (q - p) * 6 * t : t < 1 / 2 ? q : t < 2 / 3 ? p + (q - p) * (2 / 3 - t) * 6 : p; };
  return toHex([f(h + 1 / 3) * 255, f(h) * 255, f(h - 1 / 3) * 255]);
}
const relLum = (hex) => { const [r, g, b] = toRgb(hex).map((v) => { v /= 255; return v <= 0.03928 ? v / 12.92 : ((v + 0.055) / 1.055) ** 2.4; }); return 0.2126 * r + 0.7152 * g + 0.0722 * b; };
export const contrast = (a, b) => { const x = relLum(a); const y = relLum(b); return (Math.max(x, y) + 0.05) / (Math.min(x, y) + 0.05); };
const isNeutral = (hex) => { const c = toRgb(hex); return Math.max(...c) - Math.min(...c) < 30; };
/** `top` colour at `alpha` painted over `under` (both hex). */
const over = (top, alpha, under) => (alpha == null || alpha >= 1 ? top : toHex(toRgb(top).map((v, k) => v * alpha + toRgb(under)[k] * (1 - alpha))));
/** `<p:clrMapOvr><a:overrideClrMapping .../>` on a slide or layout. null = inherit (`masterClrMapping`). This is how PowerPoint stores dark background styles. */
function clrMapOf(doc) {
  const o = doc && desc(doc, 'overrideClrMapping')[0];
  if (!o) return null;
  const m = {}; for (const a of o.attributes) m[a.name] = a.value;
  return m;
}
const hidesMasterShapes = (doc) => /^(0|false)$/.test(doc?.documentElement.getAttribute('showMasterSp') || '');

/** First colour child of `holder` (solidFill, gradient stop, bgRef...). Applies lumMod/lumOff only. */
function colorOf(holder, cx) {
  const c = holder && [...holder.children].find((e) => ['srgbClr', 'schemeClr', 'sysClr', 'prstClr'].includes(e.localName));
  if (!c) return null;
  let hex = null; let slot = null;
  if (c.localName === 'srgbClr') hex = c.getAttribute('val');
  else if (c.localName === 'sysClr') hex = c.getAttribute('lastClr') || (c.getAttribute('val') === 'window' ? 'FFFFFF' : '000000');
  else if (c.localName === 'prstClr') hex = { black: '000000', white: 'FFFFFF' }[c.getAttribute('val')] || null;
  else { const v = c.getAttribute('val'); if (v === 'phClr') return null; slot = cx.clrMap[v] || v; hex = cx.theme.colors[slot]; }
  if (!/^[0-9a-f]{6}$/i.test(hex || '')) return null;
  hex = hex.toUpperCase();
  let mod = 1; let off = 0; let alpha = 1; const approx = [];
  for (const m of c.children) {
    if (m.localName === 'alpha') alpha = Math.min(1, Math.max(0, Number(m.getAttribute('val')) / 100000));
    else if (m.localName === 'lumMod') mod = Number(m.getAttribute('val')) / 100000;
    else if (m.localName === 'lumOff') off = Number(m.getAttribute('val')) / 100000;
    else if (['tint', 'shade', 'satMod', 'hueMod'].includes(m.localName)) approx.push(m.localName);
  }
  if (mod !== 1 || off) hex = lumTransform(hex, mod, off);
  return { hex, slot, alpha: alpha < 1 ? alpha : undefined, approx: approx.length ? approx : undefined };
}

// ---------- theme / master / layout ----------
async function readTheme(zip, rels) {
  const rel = [...rels.values()].find((r) => r.kind === 'theme');
  const doc = rel && (await readXml(zip, rel.target));
  if (!doc) return { ...EMPTY_THEME };
  const t = { name: doc.documentElement.getAttribute('name'), part: rel.target, colors: {}, fonts: {}, missing: false };
  const cs = desc(doc, 'clrScheme')[0];
  for (const el of cs ? cs.children : []) {
    const c = el.firstElementChild;
    const v = c && (c.localName === 'sysClr' ? c.getAttribute('lastClr') : c.getAttribute('val'));
    if (/^[0-9a-f]{6}$/i.test(v || '')) t.colors[el.localName] = v.toUpperCase();
  }
  const fs = desc(doc, 'fontScheme')[0];
  for (const [k, tag] of [['major', 'majorFont'], ['minor', 'minorFont']]) {
    const f = kid(fs, tag);
    t.fonts[k] = { latin: kid(f, 'latin')?.getAttribute('typeface') || null, ea: kid(f, 'ea')?.getAttribute('typeface') || null, cs: kid(f, 'cs')?.getAttribute('typeface') || null,
      arab: kids(f, 'font').find((x) => x.getAttribute('script') === 'Arab')?.getAttribute('typeface') || null };
  }
  return t;
}

function fillOf(holder, cx, rels, env) {
  for (const el of holder ? holder.children : []) {
    if (el.localName === 'solidFill') { const c = colorOf(el, cx); return { type: 'solid', hex: c?.hex || null, slot: c?.slot || null, alpha: c?.alpha, approx: c?.approx }; }
    if (el.localName === 'gradFill') { const stops = desc(el, 'gs').map((g) => colorOf(g, cx)?.hex).filter(Boolean); return { type: 'gradient', stops: stops.slice(0, 4), hex: null }; }
    if (el.localName === 'blipFill') { const b = desc(el, 'blip')[0]; const r = b && rels?.get(b.getAttributeNS(NS_R, 'embed')); return { type: 'image', target: r?.target || null }; }
    if (el.localName === 'noFill') return { type: 'none' };
  }
  return null;
}
function readBg(doc, cx, rels, env) {
  const bg = kid(desc(doc, 'cSld')[0], 'bg');
  if (!bg) return null;
  const pr = kid(bg, 'bgPr');
  if (pr) return fillOf(pr, cx, rels, env);
  const ref = kid(bg, 'bgRef');
  if (!ref) return null;
  const c = colorOf(ref, cx);
  return { type: 'themeRef', idx: Number(ref.getAttribute('idx')) || 0, hex: c?.hex || null, slot: c?.slot || null };
}
/** Colour a slide background reliably resolves to, or null when it can't be known (gradient, image, effect fills). */
const bgHex = (bg) => (bg?.type === 'solid' ? bg.hex : bg?.type === 'themeRef' && bg.idx <= 1001 ? bg.hex : null);

async function loadMaster(zip, path, env) {
  if (env.masters.has(path)) return env.masters.get(path);
  const doc = await readXml(zip, path);
  const rels = await readRels(zip, path);
  const theme = await readTheme(zip, rels);
  const m = { path, doc, rels, theme, clrMap: { ...DEFAULT_CLRMAP }, styles: {}, decor: [], layoutCount: 0 };
  m.cx = { theme, clrMap: m.clrMap };
  env.masters.set(path, m);
  if (!doc) return m;
  const cm = desc(doc, 'clrMap')[0];
  for (const a of cm ? cm.attributes : []) m.clrMap[a.name] = a.value;
  const tx = desc(doc, 'txStyles')[0];
  for (const n of ['titleStyle', 'bodyStyle', 'otherStyle']) m.styles[n] = kid(tx, n);
  m.decor = decorOf(doc, { ...env, cx: m.cx, rels, layoutDoc: null, master: m });
  return m;
}
async function loadLayout(zip, path, env) {
  if (env.layouts.has(path)) return env.layouts.get(path);
  const doc = await readXml(zip, path);
  const rels = await readRels(zip, path);
  const mRel = [...rels.values()].find((r) => r.kind === 'slideMaster');
  const master = mRel ? await loadMaster(zip, mRel.target, env) : null;
  const L = { path, doc, rels, master, used: 0, name: null, type: null, clrOvr: null, decor: [], phs: [] };
  env.layouts.set(path, L);
  if (!doc) return L;
  L.name = desc(doc, 'cSld')[0]?.getAttribute('name') || null;
  L.type = doc.documentElement.getAttribute('type');
  L.clrOvr = clrMapOf(doc);
  if (master) {
    L.decor = decorOf(doc, { ...env, cx: master.cx, rels, layoutDoc: null, master });
    for (const sp of desc(doc, 'sp')) {
      const ph = placeholderOf(sp);
      if (ph) L.phs.push({ type: normType(ph.type), box: xfrmOf0(sp) || xfrmOf0(phShape(ph, master.doc)) });
    }
  }
  return L;
}

const xfrmOf0 = (el) => (el ? xfrmOf(el) : null);
/** `rot` of a shape's own <a:xfrm>, in degrees. Rotation of a parent group is not read (see limitations). */
function rotOf(el) {
  const holder = el.localName === 'graphicFrame' ? el : kid(el, 'spPr');
  return (Number(kid(holder, 'xfrm')?.getAttribute('rot')) || 0) / 60000;
}
/** Axis-aligned bounds of a box rotated about its centre. A vertical rule stored as a rotated horizontal one lies where it is drawn, not where its unrotated frame is. */
const isRotated = (deg) => { const a = ((deg % 180) + 180) % 180; return a >= 0.5 && a <= 179.5; };
function rotatedBox(b, deg) {
  if (!isRotated(deg)) return b;
  const t = (((deg % 180) + 180) % 180 * Math.PI) / 180; const c = Math.abs(Math.cos(t)); const s = Math.abs(Math.sin(t));
  const w = b.w * c + b.h * s; const h = b.w * s + b.h * c;
  return { ...b, x: b.x + b.w / 2 - w / 2, y: b.y + b.h / 2 - h / 2, w, h };
}
function phShape(ph, doc) {
  if (!doc || !ph) return null;
  for (const sp of desc(doc, 'sp')) {
    const p = placeholderOf(sp);
    if (p && (ph.idx && p.idx ? ph.idx === p.idx : normType(ph.type) === normType(p.type))) return sp;
  }
  return null;
}

// ---------- text style resolution ----------
function fontName(face, theme) {
  const m = /^\+(mj|mn)-(lt|ea|cs)$/.exec(face || '');
  if (!m) return face || null;
  return theme.fonts[m[1] === 'mj' ? 'major' : 'minor']?.[{ lt: 'latin', ea: 'ea', cs: 'cs' }[m[2]]] || null;
}
function runStyle(rPr, lvl, chain, env, rtl, scale, fontCol, shapeLayers) {
  const layers = [rPr, ...chain.map((c) => kid(kids(c, `lvl${lvl}pPr`)[0], 'defRPr'))];
  const pick = (f) => { for (let i = 0; i < layers.length; i++) { const v = layers[i] && f(layers[i]); if (v != null) return [v, i]; } return [null, -1]; };
  const flag = (a) => (l) => (l.hasAttribute(a) ? /^(1|true)$/.test(l.getAttribute(a)) : null);
  const [szRaw, szI] = pick((l) => l.getAttribute('sz'));
  const [b] = pick(flag('b'));
  const [i] = pick(flag('i'));
  const [face] = pick((l) => kid(l, rtl ? 'cs' : 'latin')?.getAttribute('typeface') || null);
  // Colour precedence: run, shape list style, then the shape's <p:style><a:fontRef> colour (white text on dark autoshapes), then layout/master/defaults.
  let col = null;
  for (let k = 0; k < layers.length && !col; k++) {
    if (k === 1 + shapeLayers && fontCol) col = fontCol;
    else { const f = layers[k] && kid(layers[k], 'solidFill'); col = f ? colorOf(f, env.cx) : null; }
  }
  if (!col && fontCol && layers.length <= 1 + shapeLayers) col = fontCol;
  const theme = env.cx.theme;
  const font = fontName(face, theme) || (rtl ? theme.fonts.minor?.cs || theme.fonts.minor?.arab : null) || theme.fonts.minor?.latin || null;
  const auto = col || colorOf({ children: [{ localName: 'schemeClr', getAttribute: () => 'tx1', children: [] }] }, env.cx);
  return { sz: szRaw ? r1((Number(szRaw) / 100) * scale) : 18, src: szI === 0 ? 'explicit' : szI > 0 ? 'inherited' : 'default', b: !!b, i: !!i, font, hex: auto?.hex || null, alpha: auto?.alpha, slot: col ? col.slot : 'tx1', approx: auto?.approx };
}
function alignOf(pPr, lvl, chain) {
  const a = pPr?.getAttribute('algn') || chain.map((c) => kids(c, `lvl${lvl}pPr`)[0]?.getAttribute('algn')).find(Boolean) || 'l';
  return ALIGN[a] || 'left';
}

function readText(el, ph, lay, mas, env) {
  const body = kid(el, 'txBody');
  if (!body) return {};
  const bodyPr = kid(body, 'bodyPr');
  const scale = (Number(kid(bodyPr, 'normAutofit')?.getAttribute('fontScale')) || 100000) / 100000;
  const t = ph ? normType(ph.type) : null;
  const group = !ph || FOOTER.has(ph.type) ? 'otherStyle' : t === 'title' ? 'titleStyle' : 'bodyStyle';
  const role = !ph ? 'body' : t === 'title' ? 'title' : ph.type === 'subTitle' ? 'subtitle' : FOOTER.has(ph.type) ? 'footer' : 'body';
  const chain = [kid(body, 'lstStyle'), lay && kid(kid(lay, 'txBody'), 'lstStyle'), mas && kid(kid(mas, 'txBody'), 'lstStyle'), env.master?.styles[group], env.defaultTextStyle].filter(Boolean);
  const merged = new Map(); const lines = []; const algn = {};
  const fontRef = kid(kid(el, 'style'), 'fontRef');
  const fontCol = fontRef ? colorOf(fontRef, env.cx) : null;
  const shapeLayers = kid(body, 'lstStyle') ? 1 : 0;
  for (const p of kids(body, 'p')) {
    const pPr = kid(p, 'pPr');
    const lvl = (Number(pPr?.getAttribute('lvl')) || 0) + 1;
    let line = '';
    for (const r of [...p.children].filter((n) => n.localName === 'r' || n.localName === 'fld')) {
      const txt = kid(r, 't')?.textContent || '';
      line += txt;
      if (!txt.trim()) continue;
      const s = runStyle(kid(r, 'rPr'), lvl, chain, env, RTL.test(txt), scale, fontCol, shapeLayers);
      const key = [s.sz, s.b, s.i, s.font, s.hex, s.alpha, s.src, s.approx].join('|');
      const m = merged.get(key) || { ...s, chars: 0 };
      m.chars += txt.length;
      merged.set(key, m);
    }
    if (line.trim()) { lines.push(line); const a = alignOf(pPr, lvl, chain); algn[a] = (algn[a] || 0) + line.length; }
  }
  const text = lines.join('\n').trim();
  const ins = {}; for (const [k, a] of [['l', 'lIns'], ['t', 'tIns'], ['r', 'rIns'], ['b', 'bIns']]) if (bodyPr?.hasAttribute(a)) ins[k] = Number(bodyPr.getAttribute(a));
  return { text, words: text ? text.split(/\s+/).length : 0, runs: [...merged.values()], algn, role, insets: ins, shrunk: scale < 1 };
}

// ---------- slide shapes ----------
function recordOf(el, n, toSlide, env) {
  const ph = placeholderOf(el);
  const lay = ph && phShape(ph, env.layoutDoc);
  const mas = ph && phShape(ph, env.master?.doc);
  const own = xfrmOf(el);
  const rot = own ? rotOf(el) : 0;
  const rec = { kind: 'shape', ph: ph ? normType(ph.type) : null, phRaw: ph?.type || null, box: own ? toSlide(rotatedBox(own, rot)) : xfrmOf0(lay) || xfrmOf0(mas), rotated: isRotated(rot), fill: null, geom: null, text: '', words: 0, runs: [], image: null, textFrame: false, outline: false };
  const spPr = kid(el, 'spPr');
  if (n === 'sp') {
    rec.geom = kid(spPr, 'prstGeom')?.getAttribute('prst') || (kid(spPr, 'custGeom') ? 'custom' : null);
    rec.fill = spPr ? fillOf(spPr, env.cx, env.rels, env) : null;
    if (!rec.fill) {
      const fr = kid(kid(el, 'style'), 'fillRef');
      const c = fr && fr.getAttribute('idx') !== '0' ? colorOf(fr, env.cx) : null;
      if (c) rec.fill = { type: 'solid', hex: c.hex, slot: c.slot, approx: c.approx, viaStyle: true };
    }
    rec.textFrame = !!kid(el, 'txBody');
    rec.outline = hasOutline(el);
    Object.assign(rec, readText(el, ph, lay, mas, env));
    if (rec.words) rec.kind = 'text';
    else if (rec.fill?.type === 'image') { rec.kind = 'image'; rec.image = { target: rec.fill.target }; }
  } else if (n === 'pic') {
    rec.kind = 'image';
    const b = desc(el, 'blip')[0];
    const r = b && env.rels.get(b.getAttributeNS(NS_R, 'embed'));
    rec.image = r ? { target: r.target, external: r.external } : null;
  } else {
    rec.kind = desc(el, 'tbl').length ? 'table' : desc(el, 'chart').length ? 'chart' : 'object';
  }
  return rec;
}
/** Whether a shape draws an outline: its own <a:ln> with a fill, or one supplied by its style (<p:style><a:lnRef idx>) unless <a:ln> says noFill. */
function hasOutline(el) {
  const ln = kid(kid(el, 'spPr'), 'ln');
  if (ln) {
    if (kid(ln, 'noFill')) return false;
    if (['solidFill', 'gradFill', 'pattFill'].some((t) => kid(ln, t))) return true;
  }
  const lr = kid(kid(el, 'style'), 'lnRef');
  return !!lr && lr.getAttribute('idx') !== '0';
}
function walkTree(tree, toSlide, env, out) {
  for (const el of tree ? tree.children : []) {
    const n = el.localName;
    if (n === 'grpSp') { const g = xfrmOf(el); walkTree(el, g ? (b) => toSlide(intoGroup(b, g)) : toSlide, env, out); }
    else if (n === 'sp' || n === 'pic' || n === 'graphicFrame') out.push(recordOf(el, n, toSlide, env));
  }
}
/** Non-placeholder shapes living on a master/layout: logos, bars, frames. Part of the template's identity. */
function decorOf(doc, env) {
  const out = [];
  walkTree(desc(doc, 'spTree')[0], (b) => b, env, out);
  const { W, H } = env;
  return out.filter((s) => !s.ph && s.box && (s.kind === 'image' || s.fill?.type === 'solid' || s.kind === 'text')).slice(0, 12)
    .map((s) => ({ kind: s.kind, box: pctBox(s.box, W, H), fill: s.fill?.type === 'solid' ? s.fill.hex : null, slot: s.fill?.slot || null, approx: s.fill?.approx, text: s.kind === 'text' ? s.text.slice(0, 40) : undefined }));
}
const pctBox = (b, W, H) => ({ x: r1((b.x / W) * 100), y: r1((b.y / H) * 100), w: r1((b.w / W) * 100), h: r1((b.h / H) * 100) });

// ---------- geometry roles and backdrops (Phase 2.1) ----------
const EDGE = 0.005; // within 0.5% of the slide edge counts as touching it; overshooting by less than this is rounding, not a decision
const DEFAULT_INSET = { l: 91440, t: 45720, r: 91440, b: 45720 }; // PowerPoint's text insets (EMU) where a text box sets none
const DATA = new Set(['table', 'chart', 'object']);
const EDGES = ['left', 'right', 'top', 'bottom'];
/**
 * What a shape is for, so that decoration never distorts margins and real overflow is never hidden as decoration (Phase 2.2).
 *   footer      date / footer / slide-number placeholders
 *   empty       a placeholder nobody filled, or a text frame with no text, no fill and no outline (a "phantom" frame, e.g. the empty text box
 *               Canva-style exports leave inside every group): nothing is drawn, so it is neither decoration nor content
 *   content     everything the reader is meant to read: text, tables, charts, diagrams, and pictures/shapes inside the slide.
 *               When content runs past an edge it is overflow (`edges`, `by`): text is judged by the area its words occupy
 *               (frame minus insets: a frame that overshoots by less than its inset clips nothing), data objects by their frame
 *   background  non-text picture/panel covering (almost) the whole slide
 *   bleed       decoration that is intentionally flush with or beyond an edge: any non-text, non-data shape crossing it, or a
 *               picture touching it. Not a margin, not a problem. `by` = how far past the edge (fraction of the slide)
 *   offslide    entirely outside the slide
 * Boxes are rotation-aware (see rotatedBox), so a vertical rule stored as a rotated horizontal one is not mistaken for an off-slide shape.
 */
function roleOf(sh, W, H) {
  const b = sh.box;
  if (!b) return { role: null, edges: [], by: 0 };
  if (FOOTER.has(sh.phRaw)) return { role: 'footer', edges: [], by: 0 };
  const x0 = Math.max(0, b.x); const y0 = Math.max(0, b.y); const x1 = Math.min(W, b.x + b.w); const y1 = Math.min(H, b.y + b.h);
  if (x1 <= x0 || y1 <= y0) return { role: 'offslide', edges: [], by: 0 };
  const text = sh.words > 0;
  const bare = !text && sh.kind === 'shape' && (!sh.fill || sh.fill.type === 'none');
  if (sh.phRaw && bare) return { role: 'empty', edges: [], by: 0 };
  if (!sh.phRaw && bare && sh.textFrame && !sh.outline) return { role: 'empty', edges: [], by: 0 };
  // Area that must stay inside the slide: the text area for text, the frame for everything else.
  const ins = text && !sh.rotated ? { ...DEFAULT_INSET, ...sh.insets } : { l: 0, t: 0, r: 0, b: 0 };
  const r = { x: b.x + ins.l, y: b.y + ins.t, w: Math.max(0, b.w - ins.l - ins.r), h: Math.max(0, b.h - ins.t - ins.b) };
  const past = { left: -r.x / W, right: (r.x + r.w - W) / W, top: -r.y / H, bottom: (r.y + r.h - H) / H }; // > 0 = beyond that edge
  const crosses = Object.fromEntries(EDGES.map((k) => [k, past[k] > EDGE]));
  const reach = (edges) => Math.max(0, ...edges.map((k) => past[k]));
  if (text || DATA.has(sh.kind)) { const edges = EDGES.filter((k) => crosses[k]); return { role: 'content', edges, by: reach(edges), text }; }
  if ((x1 - x0) >= 0.95 * W && (y1 - y0) >= 0.95 * H) return { role: 'background', edges: [], by: 0 };
  const touches = { left: b.x <= EDGE * W, right: b.x + b.w >= W * (1 - EDGE), top: b.y <= EDGE * H, bottom: b.y + b.h >= H * (1 - EDGE) };
  const pic = sh.kind === 'image' || sh.fill?.type === 'image';
  const edges = EDGES.filter((k) => crosses[k] || (pic && touches[k]));
  return { role: edges.length ? 'bleed' : 'content', edges, by: reach(edges) };
}

/** Raw shape records of a master or layout under the colour map of the slide being analysed (cached per map). */
function layerRecs(part, env, cx) {
  const key = `${part.path}|${JSON.stringify(cx.clrMap)}`;
  if (!env.layerCache.has(key)) {
    const out = [];
    if (part.doc) walkTree(desc(part.doc, 'spTree')[0], (b) => b, { ...env, cx, rels: part.rels, layoutDoc: null, master: part.master || part }, out);
    env.layerCache.set(key, out);
  }
  return env.layerCache.get(key);
}

const RECTLIKE = /^(rect|roundRect|round[12]SameRect|round2DiagRect|snip[12]?Rect|snip2SameRect|snipRoundRect|flowChartProcess)$/;
const overlap = (a, b) => {
  const w = Math.min(a.x + a.w, b.x + b.w) - Math.max(a.x, b.x); const h = Math.min(a.y + a.h, b.y + b.h) - Math.max(a.y, b.y);
  return w > 0 && h > 0 && a.w > 0 && a.h > 0 ? (w * h) / (a.w * a.h) : 0;
};
/**
 * What a text shape is actually painted on: the slide background, then every shape beneath it (master, layout, earlier slide shapes,
 * its own fill) that overlaps it, composited in z-order with alpha. Returns candidate backdrop colours, or `hexes: null` with a
 * reason when the backdrop cannot be a single known colour (picture, gradient, effect fill). Partial overlaps and non-rectangular
 * shapes add a candidate instead of replacing the backdrop; contrast is only flagged when the text fails against every candidate,
 * so uncertainty never produces a false alarm.
 */
function backdropOf(sh, below, base) {
  let cand = base.hex ? [base.hex] : null;
  let reason = base.hex ? null : `${base.type === 'none' ? 'unknown' : base.type} slide background`;
  if (!sh.box) return { hexes: cand, reason };
  for (const l of [...below, sh]) {
    if (!l.box) continue;
    const cov = overlap(sh.box, l.box);
    if (cov < 0.05) continue;
    const f = l.fill; const pic = l.kind === 'image';
    if (!pic && (!f || f.type === 'none')) continue;
    if (pic || f.type === 'image' || f.type === 'gradient' || f.type !== 'solid' || !f.hex) {
      if (cov >= 0.2) { cand = null; reason = pic || f.type === 'image' ? 'picture behind the text' : 'gradient or effect fill behind the text'; }
      continue;
    }
    const a = f.alpha ?? 1;
    if (a < 0.05) continue;
    const full = l === sh || (cov >= 0.5 && (!l.geom || RECTLIKE.test(l.geom)));
    if (cand === null) { if (full && a >= 0.95) { cand = [f.hex]; reason = null; } continue; }
    const mix = cand.map((c) => over(f.hex, a, c));
    cand = [...new Set(full ? mix : [...cand, ...mix])].slice(0, 6);
  }
  return { hexes: cand, reason };
}

// ---------- classification ----------
const KW_CLOSE = /\b(thank(s| you)|questions?\??|q\s*&\s*a|get in touch|contact( us)?|the end)\b|شكر|أسئلة|اسئلة|تواصل/i;
const KW_COMPARE = /\b(vs\.?|versus|compar(e|ed|ison)|before\s*(and|&|\/)\s*after|pros\s*(and|&)\s*cons)\b|مقارنة|مقابل/i;
const KW_TITLE_NAME = /\b(title slide|cover|front page)\b/i;
const KW_CLOSE_NAME = /\b(thank|closing|end slide|the end|contact|q\s*&\s*a|goodbye)\b/i;
const TWO_COL_TYPES = ['twoObj', 'twoColTx', 'twoObjAndObj', 'twoObjAndTx', 'txAndTwoObj'];
const IMG_TYPES = ['picTx', 'txAndClipArt', 'clipArtAndTx', 'txAndMedia', 'mediaAndTx'];
const KW_IMG_NAME = /\b(pictures?|photos?|images?|gallery)\b/i;

const isTitle = (s) => s.ph === 'title' || s.titleLike;
/**
 * Slides built from plain text boxes have no title placeholder: treat the strictly-largest text box as the title. Template text (brand
 * name, slide numbers) is never a candidate. With several candidates the title must start in the top half, unless it clearly dominates
 * (at least 1.5x the next size): then it may sit lower, anywhere above the footer band, as on divider slides with the heading at the bottom.
 */
function inferTitle(shapes, H) {
  if (shapes.some((s) => s.ph === 'title')) return false;
  const c = shapes.filter((s) => s.kind === 'text' && !s.template && !FOOTER.has(s.phRaw) && s.box && s.runs.length)
    .map((s) => ({ s, sz: Math.max(...s.runs.map((r) => r.sz)) }));
  if (!c.length) return false;
  c.sort((a, b) => b.sz - a.sz);
  if (c.length === 1) { if (c[0].sz < 28 || c[0].s.box.y > H * 0.65) return false; } // a lone large text box is a title (divider slides)
  else if (c[0].sz <= c[1].sz || (c[0].s.box.y > H * 0.5 && !(c[0].sz >= 1.5 * c[1].sz && c[0].s.box.y <= H * 0.85))) return false;
  c[0].s.titleLike = true; c[0].s.role = 'title';
  return true;
}

/** A picture's identity for "repeats across the deck": same media at the same place (20 x 20 grid). */
const picKey = (s, W, H) => `${s.image.target}|${Math.round((s.box.x / W) * 20)},${Math.round((s.box.y / H) * 20)},${Math.round((s.box.w / W) * 20)},${Math.round((s.box.h / H) * 20)}`;
/** Pictures that sit at the same place on at least 30% of the slides (and 2+): logos, bands and artwork of the template, not slide content. */
function recurringPictures(S, W, H) {
  const n = S.length; const count = new Map();
  if (n < 3) return new Set();
  for (const s of S) {
    const keys = new Set(s.shapes.filter((x) => x.kind === 'image' && x.image?.target && x.box).map((x) => picKey(x, W, H)));
    keys.forEach((k) => bump(count, k));
  }
  const min = Math.max(2, Math.ceil(n * 0.3));
  return new Set([...count].filter(([, c]) => c >= min).map(([k]) => k));
}
/**
 * Template text: text that belongs to the template, not to the slide. Two kinds, both needing at least 30% of the slides (and 2+):
 *   brand       the same short text (8 words and 80 characters at most) at about the same place and size: a brand name, tagline, "Confidential"
 *   pageNumber  a plain number at about the same place whose value follows the slide order (value minus slide number is constant)
 * Shapes are only flagged (`template`), never removed: they stay in the slide structure, counts and typography. Placeholders are never flagged.
 */
const DIGITS = /^[0-9\u0660-\u0669\u06F0-\u06F9]{1,3}$/;
const digitValue = (t) => Number(t.replace(/[\u0660-\u0669]/g, (d) => d.charCodeAt(0) - 0x660).replace(/[\u06F0-\u06F9]/g, (d) => d.charCodeAt(0) - 0x6f0));
const sameSpot = (a, b, W, H) => Math.abs(a.x - b.x) <= 0.02 * W && Math.abs(a.y - b.y) <= 0.02 * H && Math.abs(a.w - b.w) <= 0.03 * W && Math.abs(a.h - b.h) <= 0.03 * H;
function markTemplateText(S, W, H) {
  const n = S.length;
  if (n < 3) return;
  const min = Math.max(2, Math.ceil(n * 0.3));
  const groups = new Map(); // key -> clusters of { box, items: [{ s, sh }] }
  for (const s of S) {
    for (const sh of s.shapes) {
      if (sh.kind !== 'text' || sh.ph || !sh.box || sh.text.length > 80 || sh.words > 8) continue; // a label, name or notice, never a sentence of body text
      const txt = sh.text.replace(/\s+/g, ' ').trim();
      const key = DIGITS.test(txt) ? '#num' : `t|${txt.toLowerCase()}`;
      const clusters = groups.get(key) || groups.set(key, []).get(key);
      const c = clusters.find((k) => sameSpot(k.box, sh.box, W, H));
      if (c) c.items.push({ s, sh, txt }); else clusters.push({ box: sh.box, items: [{ s, sh, txt }] });
    }
  }
  for (const [key, clusters] of groups) {
    for (const c of clusters) {
      const slides = new Set(c.items.map((x) => x.s.i));
      if (slides.size < min) continue;
      if (key !== '#num') { c.items.forEach((x) => { x.sh.template = 'brand'; }); continue; }
      const offsets = new Set(c.items.map((x) => digitValue(x.txt) - (x.s.i + 1)));
      if (offsets.size === 1 && slides.size === c.items.length) c.items.forEach((x) => { x.sh.template = 'pageNumber'; });
    }
  }
}

/**
 * Role of a picture that is not the slide background. `decorative`: artwork repeated across the deck, a logo or icon (under 3% of the
 * slide), or a thin band or rail along an edge (under 12% thick, over half the slide long). Everything else is `content`.
 * Only content pictures can make a slide picture-led.
 */
function pictureRole(s, W, H, deco) {
  if (!s.box) return 'content';
  const w = Math.min(W, s.box.x + s.box.w) - Math.max(0, s.box.x); const h = Math.min(H, s.box.y + s.box.h) - Math.max(0, s.box.y);
  if (s.image?.target && deco.has(picKey(s, W, H))) return 'decorative';
  if ((w * h) / (W * H) < 0.03) return 'decorative';
  if ((h <= 0.12 * H && w >= 0.5 * W) || (w <= 0.12 * W && h >= 0.5 * H)) return 'decorative';
  return 'content';
}
/**
 * Blocks that mirror each other, each with its own short heading directly above (Build | Buy, Before | After, Starter | Pro):
 * the shape of a comparison, as opposed to one text flowed into two columns. Needs 2 to 4 side-by-side blocks of similar width,
 * similar length and the same top, with headings on one line.
 */
function mirrored(blocks, W, H) {
  const headed = [];
  for (const c of blocks) {
    if (!c.box || c.words < 3 || c.box.w > 0.65 * W) continue;
    const hd = blocks.find((h) => h !== c && h.box && h.words >= 1 && h.words <= 6 && h.words < c.words
      && Math.min(h.box.x + h.box.w, c.box.x + c.box.w) - Math.max(h.box.x, c.box.x) >= 0.6 * Math.min(h.box.w, c.box.w)
      && h.box.y + h.box.h <= c.box.y + 0.02 * H && c.box.y - (h.box.y + h.box.h) <= 0.15 * H);
    if (hd) headed.push({ c, hd });
  }
  if (headed.length < 2 || headed.length > 4) return false;
  if (new Set(headed.map((x) => Math.round(x.c.box.x / (0.05 * W)))).size !== headed.length) return false; // genuinely side by side
  const within = (v, tol) => Math.max(...v) <= tol * Math.min(...v);
  const level = (v) => Math.max(...v) - Math.min(...v) <= 0.06 * H;
  return within(headed.map((x) => x.c.box.w), 1.5) && within(headed.map((x) => x.c.words), 2.5) && level(headed.map((x) => x.c.box.y)) && level(headed.map((x) => x.hd.box.y));
}

function features(shapes, bg, layout, W, H, deco = new Set()) {
  const A = W * H;
  const clipA = (s) => { if (!s.box) return 0; const w = Math.min(W, s.box.x + s.box.w) - Math.max(0, s.box.x); const h = Math.min(H, s.box.y + s.box.h) - Math.max(0, s.box.y); return w > 0 && h > 0 ? w * h : 0; };
  // A picture covering the whole slide is the slide's background, not image content: it must not turn a cover, text or closing slide into an "image" slide.
  // Of the rest, logos, icons, bands and template artwork are decoration; only the remaining pictures are content (see pictureRole).
  const pics = shapes.filter((s) => s.kind === 'image' && s.layoutRole !== 'offslide');
  const backdrops = pics.filter((s) => s.layoutRole === 'background');
  const rest = pics.filter((s) => s.layoutRole !== 'background');
  const decorPics = rest.filter((s) => pictureRole(s, W, H, deco) === 'decorative');
  const contentPics = rest.filter((s) => !decorPics.includes(s));
  const body = shapes.filter((s) => s.kind === 'text' && !FOOTER.has(s.phRaw) && !s.template); // template text (brand name, slide numbers) is not slide content
  const blocks = body.filter((s) => !isTitle(s));
  const area = (s) => (s.box ? Math.min(A, Math.max(0, s.box.w) * Math.max(0, s.box.h)) : 0);
  const cols = blocks.filter((s) => s.box && s.words >= 3);
  const side = (a, b) => {
    const ov = Math.min(a.y + a.h, b.y + b.h) - Math.max(a.y, b.y);
    return ov >= 0.6 * Math.min(a.h, b.h) && (a.x + a.w <= b.x + 0.02 * W || b.x + b.w <= a.x + 0.02 * W) && Math.max(a.w, b.w) / Math.max(1, Math.min(a.w, b.w)) <= 2 && a.w <= 0.65 * W && b.w <= 0.65 * W;
  };
  const inCols = new Set();
  cols.forEach((a, i) => cols.slice(i + 1).forEach((b) => { if (side(a.box, b.box)) { inCols.add(a); inCols.add(b); } }));
  const text = body.map((s) => s.text).join(' ');
  const title = body.find(isTitle)?.text || '';
  const L = layout?.type || null; const N = layout?.name || '';
  const words = body.reduce((n, s) => n + s.words, 0);
  // Largest font of the title against the largest of the remaining text: a cover keeps its title far above its subtitle, tagline and URL lines.
  const sizeOf = (s) => (s.runs?.length ? Math.max(...s.runs.map((r) => r.sz || 0)) : 0);
  const titleSz = sizeOf(body.find(isTitle) || { runs: [] }); const otherSz = Math.max(0, ...blocks.map(sizeOf));
  return {
    layoutType: L, layoutName: N,
    titleSz, otherSz, titleDominant: titleSz > 0 && titleSz >= 1.5 * otherSz,
    words, bodyWords: blocks.reduce((n, s) => n + s.words, 0), textBlocks: blocks.length,
    images: pics.length, contentImages: contentPics.length, decorImages: decorPics.length,
    imageCoverage: Math.min(1, pics.reduce((n, s) => n + clipA(s), 0) / A),
    contentCoverage: Math.min(1, contentPics.reduce((n, s) => n + clipA(s), 0) / A),
    backdropImage: bg?.type === 'image' || backdrops.length > 0,
    backdropTarget: backdrops[0]?.image?.target || (bg?.type === 'image' ? bg.target : null) || null,
    tables: shapes.filter((s) => s.kind === 'table').length, charts: shapes.filter((s) => s.kind === 'chart').length,
    chartArea: Math.min(1, shapes.filter((s) => s.kind === 'chart' || s.kind === 'table').reduce((n, s) => n + area(s), 0) / A),
    hasCtrTitle: shapes.some((s) => s.phRaw === 'ctrTitle'), hasTitle: shapes.some(isTitle), hasBody: shapes.some((s) => s.ph === 'body'),
    columns: inCols.size, parallel: mirrored(blocks, W, H),
    picLayout: IMG_TYPES.includes(L) || KW_IMG_NAME.test(N), twoColLayout: TWO_COL_TYPES.includes(L), compareLayout: L === 'twoTxTwoObj' || /compar/i.test(N),
    closingKw: KW_CLOSE.test(text),
    // Wording counts in the title or in a short label ("vs", "Before / After"), never in running prose ("compared with last year").
    compareTitleKw: KW_COMPARE.test(title), compareKw: KW_COMPARE.test(title) || blocks.some((s) => s.words <= 8 && KW_COMPARE.test(s.text)),
  };
}
function classify(f, i, n, cover) {
  const L = f.layoutType; const N = f.layoutName;
  const out = (c, confidence, ...reasons) => ({ class: c, confidence, reasons: reasons.filter(Boolean) });
  const pc = (v) => `${Math.round(v * 100)}%`;
  const onPic = f.backdropImage ? 'text sits on a full-slide picture, treated as the background' : null;
  if ((f.charts || f.tables) && (f.chartArea >= 0.2 || f.words <= 60 || ['tbl', 'chart', 'txAndChart', 'chartAndTx'].includes(L)))
    return out('chart-table', f.chartArea >= 0.2 ? 'high' : 'medium', f.charts ? `${f.charts} chart object(s)` : `${f.tables} table(s)`, `covers ${pc(f.chartArea)} of the slide`);
  // First slide: position and wording decide, never picture coverage (covers are often one big photo).
  const titleLayout = f.hasCtrTitle || L === 'title' || KW_TITLE_NAME.test(N);
  // More than 3 further text boxes (subtitle, tagline, URL, an un-flagged brand name) are fine on a cover when its title clearly dominates them.
  const labelled = f.textBlocks > 3 && f.titleDominant;
  if (i === 0 && (titleLayout || (f.words >= 1 && f.words <= 30 && (f.textBlocks <= 3 || f.titleDominant) && f.contentCoverage < 0.5) || (n > 1 && f.words === 0 && f.backdropImage && f.contentImages === 0)))
    return out('title', titleLayout ? 'high' : 'medium', titleLayout ? 'title layout / centred-title placeholder' : f.words ? 'first slide with little text' : 'first slide that is only a full-slide picture', 'first slide',
      !titleLayout && labelled ? `title text (${Math.round(f.titleSz)} pt) dominates the ${f.textBlocks} other text boxes (${Math.round(f.otherSz)} pt at most)` : null, onPic);
  const closeLayout = KW_CLOSE_NAME.test(N);
  if (i >= n - 2 && i > 0 && (f.closingKw || closeLayout) && f.words <= 40) return out('closing', 'high', closeLayout && !f.closingKw ? `layout name "${N}"` : 'closing wording found (thanks, questions, contact)', i === n - 1 ? 'last slide' : 'second-to-last slide', onPic);
  if (i === n - 1 && n > 2 && f.words <= 15 && (L === 'title' || f.hasCtrTitle || L === 'secHead')) return out('closing', 'medium', 'last slide using a title-style layout with little text');
  // Last slide framed like a cover: a full-slide picture, little text, nothing else.
  if (i === n - 1 && n > 2 && f.backdropImage && f.contentImages === 0 && f.words <= 25 && !f.columns && !f.tables && !f.charts) {
    const same = f.backdropTarget && f.backdropTarget === cover?.backdropTarget;
    return out('closing', same || cover?.backdropImage ? 'high' : 'medium', 'last slide with a full-slide picture and little text', same ? 'uses the same background picture as the first slide' : cover?.backdropImage ? 'mirrors the first slide\'s full-slide picture' : null, onPic);
  }
  if (i > 0 && (L === 'secHead' || /section|divider|chapter/i.test(N))) return out('section', L === 'secHead' ? 'high' : 'medium', L === 'secHead' ? 'section-header layout' : `layout name "${N}"`);
  if (i > 0 && i < n - 1 && f.hasTitle && f.words <= 10 && f.textBlocks <= 1 && f.contentImages === 0)
    return out('section', 'medium', `title with only ${f.words} words and no content pictures`, onPic);
  // Two columns are a layout choice, not a comparison. A comparison needs evidence beyond the columns: its layout, comparison
  // wording (in the title or a short label), or mirrored columns that each have their own heading. Otherwise it is text.
  const twoCol = f.columns >= 2;
  if (twoCol || (f.compareLayout && f.textBlocks >= 2)) {
    const sides = twoCol ? `${f.columns} side-by-side text blocks` : null;
    if (f.compareLayout || f.compareTitleKw) return out('comparison', 'high', f.compareLayout ? 'comparison layout' : 'comparison wording in the title', sides);
    if (f.parallel || f.compareKw) return out('comparison', 'medium', f.parallel ? 'mirrored columns, each led by its own short heading' : 'comparison wording in a short label next to side-by-side text', sides);
  }
  // Picture-led: the pictures are the content and any text only supports them. Coverage never decides alone; text density,
  // the picture's role (decorative pictures and backgrounds do not count) and the layout do. A picture beside a paragraph is an illustrated text slide.
  const plural = (n, w) => `${n} ${w}${n === 1 ? '' : 's'}`;
  if (f.contentImages >= 1) {
    const support = `${plural(f.bodyWords, 'word')} of supporting text`;
    if (f.contentCoverage >= 0.3 && f.bodyWords <= 30) return out('image', (f.contentCoverage >= 0.5 && f.bodyWords <= 15) || f.picLayout ? 'high' : 'medium', `pictures cover ${pc(f.contentCoverage)} of the slide`, support, f.picLayout ? 'picture layout' : null);
    if (f.picLayout && f.bodyWords <= 40) return out('image', 'medium', 'picture layout', support);
    if (f.contentImages >= 3 && f.contentCoverage >= 0.2 && f.bodyWords <= 40) return out('image', 'medium', `${f.contentImages} pictures covering ${pc(f.contentCoverage)} of the slide (gallery)`, support);
    if (f.contentCoverage >= 0.15 && f.bodyWords <= 15) return out('image', 'medium', `picture covering ${pc(f.contentCoverage)} of the slide with a short caption`, support);
  }
  if (f.backdropImage && f.words <= 15) return out('image', f.words === 0 ? 'high' : 'medium', f.words === 0 ? 'full-slide picture with no text' : `full-slide picture with little text (${f.words} words)`);
  if (f.words >= 1) {
    const competing = f.contentCoverage >= 0.3;
    return out('text', f.hasBody && !competing ? 'high' : 'medium', `${f.words} words`, f.hasBody ? 'body placeholder' : 'free text boxes',
      twoCol || f.twoColLayout ? 'two-column text: columns alone do not make a comparison' : null,
      f.contentImages ? `${plural(f.contentImages, 'picture')} (${pc(f.contentCoverage)} of the slide) alongside ${plural(f.bodyWords, 'word')} of text, so the text leads` : null, onPic);
  }
  return out('other', 'low', 'no readable text, pictures, tables or charts');
}

/** Final title of a slide: the inferred title text when there is one, else the parser's. Never template text such as a recurring brand name. */
function titleOf(s) {
  if (s.titleText != null) return s.titleText;
  const first = s.ps.title;
  return s.shapes.some((x) => x.template && x.text.split('\n')[0].slice(0, 160) === first) && !s.shapes.some((x) => x.ph === 'title' && x.text) ? '' : first;
}

// ---------- main ----------
/** Analyse a parsed deck (output of parsePptx, which keeps the JSZip handle). Returns the Design System JSON. */
export async function extractDesign(parsed) {
  const { zip, slides, width: W, height: H } = parsed;
  const presDoc = await readXml(zip, 'ppt/presentation.xml');
  const presRels = await readRels(zip, 'ppt/presentation.xml');
  const env = { layouts: new Map(), masters: new Map(), W, H, defaultTextStyle: desc(presDoc, 'defaultTextStyle')[0] || null,
    layerCache: new Map(), stats: { unappliedMods: new Set(), inferredTitles: 0 } };
  for (const r of presRels.values()) {
    if (r.kind !== 'slideMaster') continue;
    const m = await loadMaster(zip, r.target, env);
    env.firstMaster ||= m;
    for (const lr of m.rels.values()) if (lr.kind === 'slideLayout') { await loadLayout(zip, lr.target, env); m.layoutCount++; }
  }

  const S = [];
  for (const [i, ps] of slides.entries()) {
    if (i % 10 === 9) await tick();
    const doc = await readXml(zip, ps.part);
    const rels = await readRels(zip, ps.part);
    const lRel = [...rels.values()].find((r) => r.kind === 'slideLayout');
    const layout = lRel ? await loadLayout(zip, lRel.target, env) : null;
    if (layout) layout.used++;
    const master = layout?.master || env.firstMaster || null;
    // Colour map: the master's, replaced by a layout override, replaced by a slide override (how PowerPoint stores dark background styles).
    const slideOvr = clrMapOf(doc);
    const clrMap = { ...(master?.clrMap || DEFAULT_CLRMAP), ...(layout?.clrOvr || {}), ...(slideOvr || {}) };
    const cx = { theme: master ? master.theme : EMPTY_THEME, clrMap };
    const colorMap = slideOvr ? 'slide' : layout?.clrOvr ? 'layout' : 'master';
    const e = { ...env, cx, rels, layoutDoc: layout?.doc || null, master };
    const own = readBg(doc, cx, rels, env);
    const lbg = layout?.doc && master ? readBg(layout.doc, cx, layout.rels, env) : null;
    const mbg = master?.doc ? readBg(master.doc, cx, master.rels, env) : null;
    const bg = own ? { ...own, source: 'slide' } : lbg ? { ...lbg, source: 'layout' } : mbg ? { ...mbg, source: 'master' } : { type: 'none', source: 'default' };
    const shapes = [];
    walkTree(desc(doc, 'spTree')[0], (b) => b, e, shapes);
    const effBg = bgHex(bg) || (bg.type === 'none' && cx.theme.colors[cx.clrMap.bg1]) || null;
    // Shapes drawn beneath the slide's own shapes: master, then layout (unless hidden).
    const under = [];
    if (!hidesMasterShapes(doc)) {
      if (master && !hidesMasterShapes(layout?.doc)) under.push(...layerRecs(master, env, cx).filter((r) => !r.ph));
      if (layout) under.push(...layerRecs(layout, env, cx).filter((r) => !r.ph));
    }
    S.push({ i, ps, shapes, bg, effBg, under, f: null, cls: null, cx, layout, colorMap, titleText: null });
  }
  // Template text is only known once every slide has been read; titles, roles and backdrops are settled after that.
  markTemplateText(S, W, H);
  for (const s of S) {
    if (inferTitle(s.shapes, H)) env.stats.inferredTitles++;
    for (const sh of s.shapes) { const r = roleOf(sh, W, H); sh.layoutRole = r.role; sh.edges = r.edges; sh.by = r.by; }
    s.shapes.forEach((sh, k) => { if (sh.runs.length) sh.backdrop = backdropOf(sh, [...s.under, ...s.shapes.slice(0, k)], { hex: s.effBg, type: s.bg.type }); });
    const inferred = s.shapes.find((x) => x.titleLike);
    s.titleText = inferred ? inferred.text.split('\n')[0].slice(0, 160) : null;
  }
  // Features need the whole deck (a picture repeated at the same place on many slides is template artwork), and classification
  // needs every slide's features (the last slide is compared with the cover), so both run after the per-slide pass.
  const deco = recurringPictures(S, W, H);
  for (const s of S) s.f = features(s.shapes, s.bg, s.layout, W, H, deco);
  for (const s of S) s.cls = classify(s.f, s.i, S.length, S[0]?.f);
  return assemble(parsed, env, S);
}

function assemble(parsed, env, S) {
  const { width: W, height: H } = parsed;
  const A = W * H; const n = S.length;
  const theme = S[0]?.cx.theme || env.firstMaster?.theme || EMPTY_THEME;
  const themeByHex = new Map(Object.entries(theme.colors).map(([k, v]) => [v, k]));
  const limitations = [
    'Slide classes are rule-based heuristics. Each slide shows its confidence and the reasons; treat "low" and "medium" as suggestions.',
    'PPTX stores weight only as bold or regular; named weights (e.g. "Semibold") appear inside the font name.',
    'Charts, tables and embedded objects are detected by presence, position and size only; their internal styling is not analysed.',
    'Picture pixel resolution and picture content are not analysed.',
  ];
  if (theme.missing) limitations.push('No theme part was found, so theme colours and fonts are unavailable and scheme-based colours could not be resolved.');
  if (env.stats.unappliedMods.size) limitations.push(`Colour transforms not applied: ${[...env.stats.unappliedMods].join(', ')}. Affected colours may be approximate.`);
  if (env.stats.inferredTitles) limitations.push(`${env.stats.inferredTitles} slide(s) have no title placeholder; the largest text box was treated as the title.`);
  limitations.push('Rotation of a single shape is applied to its bounds; rotation and flips of groups are ignored when positioning shapes.');
  limitations.push('Decorative bleed (pictures and shapes flush with or past an edge) is told apart from content overflow by shape type and position only; text is judged by its frame minus insets, so text that wraps beyond its frame is not detected.');

  // typography, palette, alignment, shapes
  const fonts = new Map(); const sizes = new Map(); const roles = {}; const pal = new Map();
  const algn = {}; let totChars = 0; let boldChars = 0; let unresolved = 0; let totFill = 0; let shrunk = 0;
  const kinds = {}; const presets = {}; const insets = { l: [], t: [], r: [], b: [] };
  const addPal = (c, k, w) => { if (!c?.hex) return; (c.approx || []).forEach((m) => env.stats.unappliedMods.add(m)); const e = pal.get(c.hex) || { hex: c.hex, slot: null, text: 0, fill: 0, bg: 0 }; e[k] += w; e.slot ||= c.slot && !/^(tx|bg)/.test(c.slot) ? c.slot : null; pal.set(c.hex, e); };
  const bgTypes = {}; const bgColors = new Map(); const slideRecs = [];
  const unevalBg = { n: 0 };

  for (const s of S) {
    for (const d of [...(s.layout?.decor || []), ...(s.layout?.master?.decor || [])]) {
      if (d.fill && d.box.w * d.box.h < 8100) { const w = (d.box.w * d.box.h) / 10000; addPal({ hex: d.fill, slot: d.slot, approx: d.approx }, 'fill', w); totFill += w; }
    }
    // Margins describe the content frame only: pictures and panels that bleed off the slide, full-slide backgrounds and template text are decoration.
    const content = s.shapes.filter((x) => x.layoutRole === 'content' && !x.template && ['text', 'image', 'table', 'chart', 'object'].includes(x.kind) && !(x.box.w >= 0.9 * W && x.box.h >= 0.9 * H));
    // Decorative bleed is intentional (reported, never a problem); overflow is content (text, table, chart, diagram) that runs past
    // the slide edge or sits entirely outside it. The two are reported separately.
    const bleedShapes = s.shapes.filter((x) => x.layoutRole === 'bleed');
    const bleed = [...new Set(bleedShapes.flatMap((x) => x.edges))];
    const bleedDetail = bleedShapes.slice(0, 4).map((x) => ({ kind: x.kind === 'image' ? 'picture' : 'shape', edges: x.edges, mode: x.by > EDGE ? 'beyond' : 'flush', byPct: r1(x.by * 100), box: x.box }));
    const overflow = s.shapes.filter((x) => (x.words > 0 || DATA.has(x.kind)) && (x.layoutRole === 'offslide' || (x.layoutRole === 'content' && x.edges.length)))
      .map((x) => ({ kind: x.kind, edges: x.edges, byPct: r1(x.by * 100), outside: x.layoutRole === 'offslide' || undefined, sample: x.words ? x.text.slice(0, 30) : undefined }));
    const offSlide = overflow.length > 0;
    let checked = 0; let unchecked = 0;
    let minPt = null; let maxPt = null; let titleSz = null; let smallChars = 0; let worst = null; let shapeCount = {};
    for (const sh of s.shapes) {
      kinds[sh.kind] = (kinds[sh.kind] || 0) + 1; shapeCount[sh.kind] = (shapeCount[sh.kind] || 0) + 1;
      if (sh.geom) presets[sh.geom] = (presets[sh.geom] || 0) + 1;
      if (sh.fill?.type === 'solid' && sh.box) { const w = Math.min(1, (sh.box.w * sh.box.h) / A); if (w < 0.9) { addPal(sh.fill, 'fill', w); totFill += w; } }
      for (const k of ['l', 't', 'r', 'b']) if (sh.insets?.[k] != null) insets[k].push(sh.insets[k]);
      for (const [a, c] of Object.entries(sh.algn || {})) if (sh.role !== 'footer') algn[a] = (algn[a] || 0) + c;
      if (sh.shrunk) shrunk++;
      for (const r of sh.runs) {
        totChars += r.chars; if (r.b) boldChars += r.chars; if (r.src === 'default') unresolved += r.chars;
        bump(fonts, r.font || 'Unknown', r.chars); bump(sizes, String(Math.round(r.sz * 2) / 2), r.chars);
        addPal({ hex: r.hex, slot: r.slot, approx: r.approx }, 'text', r.chars);
        const ro = (roles[sh.role] ||= { font: new Map(), sz: new Map(), b: new Map(), color: new Map(), chars: 0 });
        bump(ro.font, r.font, r.chars); bump(ro.sz, r.sz, r.chars); bump(ro.b, r.b, r.chars); bump(ro.color, r.hex, r.chars); ro.chars += r.chars;
        if (sh.role !== 'footer') {
          minPt = minPt == null ? r.sz : Math.min(minPt, r.sz); maxPt = maxPt == null ? r.sz : Math.max(maxPt, r.sz);
          if (r.sz < 12) smallChars += r.chars;
        }
        if (sh.role === 'title') titleSz = r.sz;
        // Contrast is judged against what the text is really painted on (see backdropOf), with text alpha applied.
        const bd = sh.backdrop;
        if (!bd?.hexes) { unevalBg.n += r.chars; unchecked += r.chars; }
        else if (r.hex) {
          checked += r.chars;
          let best = null;
          for (const c of bd.hexes) { const t = over(r.hex, r.alpha, c); const ratio = contrast(t, c); if (!best || ratio > best.ratio) best = { ratio, text: t, background: c }; }
          const large = r.sz >= 18 || (r.b && r.sz >= 14);
          if (best.ratio < (large ? 3 : 4.5) && (!worst || best.ratio < worst.ratio)) worst = { ratio: r2(best.ratio), text: best.text, background: best.background, sizePt: r.sz, sample: sh.text.slice(0, 30) };
        }
      }
    }
    bgTypes[s.bg.type] = (bgTypes[s.bg.type] || 0) + 1;
    if (s.bg.hex) { addPal(s.bg, 'bg', 1); bump(bgColors, s.bg.hex, 1); } else if (s.effBg && s.bg.type === 'none') { addPal({ hex: s.effBg, slot: 'bg1' }, 'bg', 1); bump(bgColors, s.effBg, 1); }

    // margins from content bounds
    let bounds = null;
    if (content.length) {
      const x0 = Math.max(0, Math.min(...content.map((c) => c.box.x))); const y0 = Math.max(0, Math.min(...content.map((c) => c.box.y)));
      const x1 = Math.min(W, Math.max(...content.map((c) => c.box.x + c.box.w))); const y1 = Math.min(H, Math.max(...content.map((c) => c.box.y + c.box.h)));
      bounds = { left: r1((x0 / W) * 100), top: r1((y0 / H) * 100), right: r1(((W - x1) / W) * 100), bottom: r1(((H - y1) / H) * 100) };
    }
    // vertical gaps between stacked content
    const gaps = []; const stack = content.filter((c) => c.kind !== 'object').sort((a, b) => a.box.y - b.box.y);
    for (let k = 1; k < stack.length; k++) {
      const a = stack[k - 1].box; const b = stack[k].box;
      const ov = Math.min(a.x + a.w, b.x + b.w) - Math.max(a.x, b.x);
      if (ov >= 0.5 * Math.min(a.w, b.w) && b.y >= a.y + a.h) gaps.push(((b.y - a.y - a.h) / H) * 100);
    }
    Object.assign(s, { bounds, minPt, maxPt, titleSz, smallChars, worst, gaps, shapeCount, content, bleed, bleedDetail, overflow: overflow.slice(0, 4), offSlide, checked, unchecked });
  }

  // palette
  const totText = totChars || 0; const totBg = n;
  const dims = [totText, totFill, totBg].filter(Boolean).length || 1;
  const palette = [...pal.values()].map((e) => {
    const roleList = [e.text && 'text', e.fill && 'fill', e.bg && 'background'].filter(Boolean);
    return { hex: e.hex, slot: e.slot || themeByHex.get(e.hex) || null, neutral: isNeutral(e.hex), roles: roleList,
      weight: r2(((totText ? e.text / totText : 0) + (totFill ? e.fill / totFill : 0) + (totBg ? e.bg / totBg : 0)) / dims) };
  }).sort((a, b) => b.weight - a.weight);
  const offTheme = palette.filter((c) => !c.slot && !c.neutral);

  // typography
  const share = (map) => [...map.entries()].map(([k, v]) => [k, v / (totText || 1)]).sort((a, b) => b[1] - a[1]);
  const roleOut = {};
  for (const [k, ro] of Object.entries(roles)) roleOut[k] = { font: topKey(ro.font), sizePt: Number(topKey(ro.sz)), bold: !!topKey(ro.b), color: topKey(ro.color), chars: ro.chars };
  const sizeList = share(sizes).map(([pt, sh]) => ({ pt: Number(pt), share: r2(sh) })).sort((a, b) => a.pt - b.pt);
  const typography = {
    fonts: share(fonts).filter(([, sh]) => sh > 0).slice(0, 8).map(([name, sh]) => ({ name, share: r2(sh) })),
    sizes: sizeList.slice(0, 16), roles: roleOut,
    minPt: sizeList.length ? sizeList[0].pt : null, maxPt: sizeList.length ? sizeList[sizeList.length - 1].pt : null,
    boldShare: r2(boldChars / (totText || 1)), characters: totText,
  };
  if (unresolved) limitations.push(`${Math.round((unresolved / totText) * 100)}% of text had no size anywhere in the file's style chain; 18 pt (the PowerPoint default) was assumed for it.`);
  if (shrunk) limitations.push(`${shrunk} text box(es) use "shrink text on overflow"; sizes were scaled by the stored factor, which PowerPoint may recompute when opened.`);
  if (unevalBg.n && totText) limitations.push(`Contrast was not evaluated for ${Math.round((unevalBg.n / totText) * 100)}% of text (it sits on a picture, gradient or effect fill, which has no single colour).`);
  if (!totText) limitations.push('This template has no text, so typography and contrast could not be analysed.');

  // spacing
  const lists = { left: [], right: [], top: [], bottom: [] };
  for (const s of S) if (s.bounds) for (const k of Object.keys(lists)) lists[k].push(s.bounds[k]);
  const margins = Object.fromEntries(Object.entries(lists).map(([k, v]) => [k, r1(median(v))]));
  const spread = Object.fromEntries(Object.entries(lists).map(([k, v]) => [k, v.length ? r1(pctile(v, 0.9) - pctile(v, 0.1)) : null]));
  const allGaps = S.flatMap((s) => s.gaps);
  const totAlign = Object.values(algn).reduce((a, b) => a + b, 0) || 1;
  const spacing = {
    marginsPct: margins, spreadPct: spread, gapPct: r1(median(allGaps)),
    bleedSlides: S.filter((s) => s.bleed.length).length, offSlideSlides: S.filter((s) => s.offSlide).length,
    textInsetsEmu: Object.fromEntries(Object.entries(insets).map(([k, v]) => [k, v.length ? median(v) : { l: 91440, t: 45720, r: 91440, b: 45720 }[k]])),
    insetsSource: insets.l.length || insets.t.length ? 'explicit where set, PowerPoint defaults elsewhere' : 'PowerPoint defaults',
    alignment: Object.fromEntries(Object.entries(algn).map(([k, v]) => [k, r2(v / totAlign)])),
  };

  // patterns and recurring elements
  const bin = (b) => `${Math.round((b.x / W) * 20)},${Math.round((b.y / H) * 20)},${Math.round((b.w / W) * 20)},${Math.round((b.h / H) * 20)}`;
  // Decorative bleed that sits at the same place on several slides is a template element (bar, band, glow): it is why many slides share a bleed flag.
  const bleedSig = (x) => `${x.kind}|${x.edges.join()}|${bin(x.box)}`;
  const bleedSeen = new Map();
  for (const s of S) new Set(s.bleedDetail.map(bleedSig)).forEach((k) => bump(bleedSeen, k));
  for (const s of S) for (const x of s.bleedDetail) { const c = bleedSeen.get(bleedSig(x)); if (c >= 2) x.repeats = c; delete x.box; }
  const sigs = new Map(); const recur = new Map();
  for (const s of S) {
    const sig = `${s.layout?.name || '-'}|` + s.shapes.filter((x) => x.box && ['text', 'image', 'table', 'chart'].includes(x.kind)).map((x) => `${x.kind}${x.ph || ''}@${bin(x.box)}`).sort().join(';');
    (sigs.get(sig) || sigs.set(sig, []).get(sig)).push(s.i + 1);
    const seen = new Set();
    for (const x of s.shapes) {
      if (!x.box) continue;
      const key = x.kind === 'image' && x.image?.target ? `image|${x.image.target}|${bin(x.box)}` : x.kind === 'text' && x.text.length <= 80 && FOOTER.has(x.phRaw) === false && !x.ph ? `text|${x.text}|${bin(x.box)}` : null;
      if (key && !seen.has(key)) { seen.add(key); const e = recur.get(key) || { kind: x.kind, label: x.kind === 'text' ? x.text : x.image.target.replace(/^ppt\//, ''), box: pctBox(x.box, W, H), slides: [] }; e.slides.push(s.i + 1); recur.set(key, e); }
    }
  }
  const patterns = [...sigs.entries()].filter(([, v]) => v.length >= 2).sort((a, b) => b[1].length - a[1].length).slice(0, 12)
    .map(([sig, v], k) => { const cm = new Map(); v.forEach((x) => bump(cm, S[x - 1].cls.class)); return { id: `p${k + 1}`, layout: sig.split('|')[0] === '-' ? null : sig.split('|')[0], slides: v, class: topKey(cm) }; });
  const minRep = Math.max(2, Math.ceil(n * 0.3));
  const recurring = [...recur.values()].filter((e) => e.slides.length >= minRep).sort((a, b) => b.slides.length - a.slides.length).slice(0, 10).map((e) => ({ ...e, count: e.slides.length, slides: undefined }));

  // layouts
  const layouts = { masters: [...env.masters.values()].map((m) => ({ part: m.path, theme: m.theme.name, layouts: m.layoutCount, decor: m.decor })),
    list: [...env.layouts.values()].filter((L) => L.doc).slice(0, 40).map((L) => ({ name: L.name, type: L.type, used: L.used, decor: L.decor.length, placeholders: L.phs.filter((p) => p.box).slice(0, 8).map((p) => ({ type: p.type, box: pctBox(p.box, W, H) })) })) };

  // per-slide records + classes
  const classes = Object.fromEntries(CLASSES.map((c) => [c, 0]));
  const conf = { high: 0, medium: 0, low: 0 };
  for (const s of S) {
    classes[s.cls.class]++; conf[s.cls.confidence]++;
    slideRecs.push({ n: s.i + 1, title: titleOf(s), templateText: s.shapes.filter((x) => x.template).map((x) => ({ kind: x.template, text: x.text.slice(0, 40) })), layout: s.layout?.name || null, layoutType: s.layout?.type || null, class: s.cls.class, confidence: s.cls.confidence, reasons: s.cls.reasons,
      background: { type: s.bg.type, hex: s.bg.hex || s.effBg || null, source: s.bg.source, picture: s.f.backdropImage || undefined }, words: s.f.words, counts: s.shapeCount,
      imageCoverage: r2(s.f.imageCoverage), contentImageCoverage: r2(s.f.contentCoverage), imageRoles: s.f.images ? { background: s.f.images - s.f.contentImages - s.f.decorImages, content: s.f.contentImages, decorative: s.f.decorImages } : undefined,
      bounds: s.bounds, bleed: s.bleed.length ? s.bleed : undefined, bleedDetail: s.bleedDetail.length ? s.bleedDetail : undefined, offSlide: s.offSlide || undefined, overflow: s.overflow.length ? s.overflow : undefined, colorMap: s.colorMap, minPt: s.minPt, maxPt: s.maxPt,
      contrast: { status: s.worst ? 'low' : s.checked ? 'ok' : s.unchecked ? 'not-evaluated' : 'no-text', ratio: s.worst?.ratio, text: s.worst?.text, background: s.worst?.background, unchecked: s.unchecked ? r2(s.unchecked / (s.checked + s.unchecked)) : undefined } });
  }

  const bgDominant = topKey(bgColors);
  const modes = buildModes({ typography, spacing, palette, theme, bgDominant, S, n, bgColors });
  const design = {
    schema: 'slideflow.design', version: DESIGN_VERSION, analyzedAt: Date.now(),
    slide: { width: W, height: H, aspect: r2(W / H) },
    theme: { name: theme.name, part: theme.part, colors: theme.colors, fonts: { major: theme.fonts.major?.latin || null, minor: theme.fonts.minor?.latin || null } },
    palette: { colors: palette.slice(0, 14), offThemeCount: offTheme.length },
    typography, backgrounds: { types: bgTypes, colors: [...bgColors.entries()].map(([hex, slides]) => ({ hex, slides })).sort((a, b) => b.slides - a.slides).slice(0, 6) },
    shapes: { total: Object.values(kinds).reduce((a, b) => a + b, 0), byKind: kinds, presets },
    layouts, spacing, patterns, recurring,
    classification: { counts: classes, confidence: conf }, slides: slideRecs, modes, limitations,
  };
  const clean = JSON.parse(JSON.stringify(design)); // drops undefined/NaN so it is storable in Firestore (no nested arrays are produced)
  if (JSON.stringify(clean).length > 700000) clean.slides.forEach((s) => { delete s.reasons; });
  return clean;
}

// ---------- modes ----------
function buildModes({ typography: t, spacing: sp, palette, theme, bgDominant, S, n }) {
  const r = t.roles;
  const firstFont = t.fonts[0]?.name || null;
  const observed = {
    fonts: { heading: r.title?.font || firstFont, body: r.body?.font || firstFont },
    sizesPt: { title: r.title?.sizePt ?? null, subtitle: r.subtitle?.sizePt ?? null, body: r.body?.sizePt ?? null, footer: r.footer?.sizePt ?? null },
    marginsPct: { ...sp.marginsPct },
    colors: { text: r.body?.color || r.title?.color || null, background: bgDominant || null, accents: palette.filter((c) => !c.neutral).slice(0, 4).map((c) => c.hex) },
  };
  const enh = JSON.parse(JSON.stringify(observed));
  const changes = [];
  const set = (path, to, reason) => {
    const parts = path.split('.'); let o = observed; let e = enh;
    for (const p of parts.slice(0, -1)) { o = o[p]; e = e[p]; }
    const k = parts[parts.length - 1];
    if (to != null && o[k] !== to) { changes.push({ token: path, from: o[k], to, reason }); e[k] = to; }
  };
  const body = observed.sizesPt.body;
  if (body != null) set('sizesPt.body', Math.max(body, 16), 'Body text should be at least 16 pt to stay readable when projected.');
  const eb = enh.sizesPt.body;
  if (observed.sizesPt.title != null) set('sizesPt.title', Math.max(observed.sizesPt.title, 28, eb != null ? Math.round(eb * 1.25) : 0), 'Titles should stay clearly larger than body text (at least 28 pt and 1.25x body).');
  if (observed.sizesPt.subtitle != null && eb != null) set('sizesPt.subtitle', Math.max(observed.sizesPt.subtitle, eb), 'A subtitle should never be smaller than body text.');
  if (observed.sizesPt.footer != null) set('sizesPt.footer', Math.max(observed.sizesPt.footer, 12), 'Footers and slide numbers stay at 12 pt or more.');
  const floor = { left: 5, right: 5, top: 5, bottom: 4 };
  for (const k of Object.keys(floor)) if (observed.marginsPct[k] != null) set(`marginsPct.${k}`, r1(Math.max(observed.marginsPct[k], floor[k])), 'Outer margins are kept at the template median and never below a safe minimum, so every slide shares one frame.');
  if (observed.colors.text && observed.colors.background && contrast(observed.colors.text, observed.colors.background) < 4.5) {
    const best = [theme.colors.dk1, theme.colors.lt1].filter(Boolean).sort((a, b) => contrast(b, observed.colors.background) - contrast(a, observed.colors.background))[0];
    if (best) set('colors.text', best, 'Body text colour has under 4.5:1 contrast with the dominant background; the theme colour with the best contrast is used.');
  }

  const findings = [];
  const add = (code, severity, message, slides = [], evidence) => findings.push({ code, severity, message, slides: slides.slice(0, 20), evidence });
  const small = S.filter((s) => s.smallChars > 0).map((s) => s.i + 1);
  if (small.length) add('small-text', 'warn', 'Text under 12 pt is hard to read on a projector.', small);
  if (t.sizes.length && body != null && body < 14) add('low-body-size', 'info', `The dominant body size is ${body} pt; 16 pt or more reads better on screen.`);
  const manyFonts = t.fonts.filter((f) => f.share >= 0.02);
  if (manyFonts.length > 2) add('many-fonts', 'warn', `${manyFonts.length} font families carry more than 2% of the text each (${manyFonts.map((f) => f.name).join(', ')}). Two is the usual limit.`);
  const off = palette.filter((c) => !c.slot && !c.neutral);
  if (off.length >= 3) add('off-theme-colors', 'info', `${off.length} hard-coded colours are not part of the theme palette, so they will not follow a theme change.`);
  const lowS = S.filter((s) => s.worst); const low = lowS.map((s) => s.i + 1);
  if (low.length) add('low-contrast', 'warn', 'Text/background contrast is below WCAG AA (4.5:1 for normal text, 3:1 for large text).', low,
    lowS.slice(0, 5).map((s) => ({ slide: s.i + 1, ratio: s.worst.ratio, text: s.worst.text, background: s.worst.background, sizePt: s.worst.sizePt, sample: s.worst.sample })));
  const lefts = S.filter((s) => s.bounds).map((s) => s.bounds.left); const medL = median(lefts);
  const offMargin = S.filter((s) => s.bounds && Math.abs(s.bounds.left - medL) > 4).map((s) => s.i + 1);
  if (offMargin.length >= 2 && offMargin.length < n) add('inconsistent-margins', 'info', 'Left margins differ by more than 4% of the slide width from the template median.', offMargin);
  const tsz = S.filter((s) => s.cls.class === 'text' && s.titleSz != null); const medT = median(tsz.map((s) => s.titleSz));
  const oddT = tsz.filter((s) => Math.abs(s.titleSz - medT) > 4).map((s) => s.i + 1);
  if (oddT.length) add('inconsistent-title-size', 'info', `Titles on text slides vary by more than 4 pt from the typical ${medT} pt.`, oddT);
  const dense = S.filter((s) => s.f.words > 110).map((s) => s.i + 1);
  if (dense.length) add('dense-text', 'info', 'More than 110 words on one slide.', dense);
  const over = S.filter((s) => s.offSlide);
  if (over.length) add('content-overflow', 'warn', 'Text, a table, a chart or a diagram runs past the slide edge. Decorative pictures and shapes that bleed off the edge on purpose are not counted.', over.map((s) => s.i + 1),
    over.slice(0, 5).map((s) => ({ slide: s.i + 1, kind: s.overflow[0].kind, edges: s.overflow[0].edges, byPct: s.overflow[0].byPct, sample: s.overflow[0].sample })));

  return {
    preserve: { description: 'Keep the template exactly as analysed. Tokens are the observed fonts, sizes, colours and margins; nothing is normalised.', tokens: observed },
    enhance: { description: 'Keep the template\'s fonts, palette and layout identity, but raise sizes, margins and contrast to readable minimums. Only values listed under "changes" differ from Preserve.', tokens: enh, changes, findings },
    note: 'Modes are a stored contract for slide generation (Phase 3). Phase 2 never modifies your PPTX file.',
  };
}
