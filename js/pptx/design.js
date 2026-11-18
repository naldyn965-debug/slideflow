// Template intelligence (Phase 2): theme, text styles, backgrounds, layout/spacing statistics, slide classification
// and the JSON "Design System". Read-only and rule-based: every number comes from the file's XML, every class from
// stated rules with a confidence. Anything that cannot be resolved is reported in `limitations`, never guessed.
import { kid, kids, desc, readXml, readRels, xfrmOf, intoGroup, placeholderOf, normType } from './parser.js';

export const DESIGN_VERSION = 1;
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
  let mod = 1; let off = 0; const approx = [];
  for (const m of c.children) {
    if (m.localName === 'lumMod') mod = Number(m.getAttribute('val')) / 100000;
    else if (m.localName === 'lumOff') off = Number(m.getAttribute('val')) / 100000;
    else if (['tint', 'shade', 'satMod', 'hueMod'].includes(m.localName)) approx.push(m.localName);
  }
  if (mod !== 1 || off) hex = lumTransform(hex, mod, off);
  return { hex, slot, approx: approx.length ? approx : undefined };
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
    if (el.localName === 'solidFill') { const c = colorOf(el, cx); return { type: 'solid', hex: c?.hex || null, slot: c?.slot || null, approx: c?.approx }; }
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
  const m = { path, doc, rels, theme, clrMap: { ...DEFAULT_CLRMAP }, styles: {}, bg: null, decor: [], layoutCount: 0 };
  m.cx = { theme, clrMap: m.clrMap };
  env.masters.set(path, m);
  if (!doc) return m;
  const cm = desc(doc, 'clrMap')[0];
  for (const a of cm ? cm.attributes : []) m.clrMap[a.name] = a.value;
  const tx = desc(doc, 'txStyles')[0];
  for (const n of ['titleStyle', 'bodyStyle', 'otherStyle']) m.styles[n] = kid(tx, n);
  m.bg = readBg(doc, m.cx, rels, env);
  m.decor = decorOf(doc, { ...env, cx: m.cx, rels, layoutDoc: null, master: m });
  return m;
}
async function loadLayout(zip, path, env) {
  if (env.layouts.has(path)) return env.layouts.get(path);
  const doc = await readXml(zip, path);
  const rels = await readRels(zip, path);
  const mRel = [...rels.values()].find((r) => r.kind === 'slideMaster');
  const master = mRel ? await loadMaster(zip, mRel.target, env) : null;
  const L = { path, doc, rels, master, used: 0, name: null, type: null, bg: null, decor: [], phs: [] };
  env.layouts.set(path, L);
  if (!doc) return L;
  L.name = desc(doc, 'cSld')[0]?.getAttribute('name') || null;
  L.type = doc.documentElement.getAttribute('type');
  if (desc(doc, 'overrideClrMapping').length) env.stats.clrOverride = true;
  if (master) {
    L.bg = readBg(doc, master.cx, rels, env);
    L.decor = decorOf(doc, { ...env, cx: master.cx, rels, layoutDoc: null, master });
    for (const sp of desc(doc, 'sp')) {
      const ph = placeholderOf(sp);
      if (ph) L.phs.push({ type: normType(ph.type), box: xfrmOf0(sp) || xfrmOf0(phShape(ph, master.doc)) });
    }
  }
  return L;
}

const xfrmOf0 = (el) => (el ? xfrmOf(el) : null);
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
function runStyle(rPr, lvl, chain, env, rtl, scale) {
  const layers = [rPr, ...chain.map((c) => kid(kids(c, `lvl${lvl}pPr`)[0], 'defRPr'))];
  const pick = (f) => { for (let i = 0; i < layers.length; i++) { const v = layers[i] && f(layers[i]); if (v != null) return [v, i]; } return [null, -1]; };
  const flag = (a) => (l) => (l.hasAttribute(a) ? /^(1|true)$/.test(l.getAttribute(a)) : null);
  const [szRaw, szI] = pick((l) => l.getAttribute('sz'));
  const [b] = pick(flag('b'));
  const [i] = pick(flag('i'));
  const [face] = pick((l) => kid(l, rtl ? 'cs' : 'latin')?.getAttribute('typeface') || null);
  const [col] = pick((l) => { const f = kid(l, 'solidFill'); return f ? colorOf(f, env.cx) : null; });
  const theme = env.cx.theme;
  const font = fontName(face, theme) || (rtl ? theme.fonts.minor?.cs || theme.fonts.minor?.arab : null) || theme.fonts.minor?.latin || null;
  const auto = col || colorOf({ children: [{ localName: 'schemeClr', getAttribute: () => 'tx1', children: [] }] }, env.cx);
  return { sz: szRaw ? r1((Number(szRaw) / 100) * scale) : 18, src: szI === 0 ? 'explicit' : szI > 0 ? 'inherited' : 'default', b: !!b, i: !!i, font, hex: auto?.hex || null, slot: col ? col.slot : 'tx1', approx: auto?.approx };
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
  for (const p of kids(body, 'p')) {
    const pPr = kid(p, 'pPr');
    const lvl = (Number(pPr?.getAttribute('lvl')) || 0) + 1;
    let line = '';
    for (const r of [...p.children].filter((n) => n.localName === 'r' || n.localName === 'fld')) {
      const txt = kid(r, 't')?.textContent || '';
      line += txt;
      if (!txt.trim()) continue;
      const s = runStyle(kid(r, 'rPr'), lvl, chain, env, RTL.test(txt), scale);
      const key = [s.sz, s.b, s.i, s.font, s.hex, s.src, s.approx].join('|');
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
  const rec = { kind: 'shape', ph: ph ? normType(ph.type) : null, phRaw: ph?.type || null, box: own ? toSlide(own) : xfrmOf0(lay) || xfrmOf0(mas), fill: null, geom: null, text: '', words: 0, runs: [], image: null };
  const spPr = kid(el, 'spPr');
  if (n === 'sp') {
    rec.geom = kid(spPr, 'prstGeom')?.getAttribute('prst') || (kid(spPr, 'custGeom') ? 'custom' : null);
    rec.fill = spPr ? fillOf(spPr, env.cx, env.rels, env) : null;
    if (!rec.fill) {
      const fr = kid(kid(el, 'style'), 'fillRef');
      const c = fr && fr.getAttribute('idx') !== '0' ? colorOf(fr, env.cx) : null;
      if (c) rec.fill = { type: 'solid', hex: c.hex, slot: c.slot, approx: c.approx, viaStyle: true };
    }
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

// ---------- classification ----------
const KW_CLOSE = /\b(thank(s| you)|questions?\??|q\s*&\s*a|get in touch|contact( us)?|the end)\b|شكر|أسئلة|اسئلة|تواصل/i;
const KW_COMPARE = /\b(vs\.?|versus|compar(e|ed|ison)|before\s*(and|&|\/)\s*after|pros\s*(and|&)\s*cons)\b|مقارنة|مقابل/i;
const TWO_COL_TYPES = ['twoObj', 'twoColTx', 'twoObjAndObj', 'twoObjAndTx', 'txAndTwoObj'];
const IMG_TYPES = ['picTx', 'txAndClipArt', 'clipArtAndTx', 'txAndMedia', 'mediaAndTx'];

const isTitle = (s) => s.ph === 'title' || s.titleLike;
/** Slides built from plain text boxes have no title placeholder: treat the top, strictly-largest text box as the title. */
function inferTitle(shapes, H) {
  if (shapes.some((s) => s.ph === 'title')) return false;
  const c = shapes.filter((s) => s.kind === 'text' && !FOOTER.has(s.phRaw) && s.box && s.runs.length)
    .map((s) => ({ s, sz: Math.max(...s.runs.map((r) => r.sz)) }));
  if (c.length < 2) return false;
  c.sort((a, b) => b.sz - a.sz);
  if (c[0].sz <= c[1].sz || c[0].s.box.y > H * 0.5) return false;
  c[0].s.titleLike = true; c[0].s.role = 'title';
  return true;
}

function features(shapes, bg, layout, W, H) {
  const A = W * H;
  const body = shapes.filter((s) => s.kind === 'text' && !FOOTER.has(s.phRaw));
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
  return {
    layoutType: layout?.type || null, layoutName: layout?.name || '',
    words: body.reduce((n, s) => n + s.words, 0), textBlocks: blocks.length,
    images: shapes.filter((s) => s.kind === 'image').length,
    imageCoverage: Math.min(1, shapes.filter((s) => s.kind === 'image').reduce((n, s) => n + area(s), 0) / A),
    bgImage: bg?.type === 'image',
    tables: shapes.filter((s) => s.kind === 'table').length, charts: shapes.filter((s) => s.kind === 'chart').length,
    chartArea: Math.min(1, shapes.filter((s) => s.kind === 'chart' || s.kind === 'table').reduce((n, s) => n + area(s), 0) / A),
    hasCtrTitle: shapes.some((s) => s.phRaw === 'ctrTitle'), hasTitle: shapes.some(isTitle), hasBody: shapes.some((s) => s.ph === 'body'),
    columns: inCols.size, closingKw: KW_CLOSE.test(text), compareKw: KW_COMPARE.test(title) || KW_COMPARE.test(text.slice(0, 200)),
  };
}
function classify(f, i, n) {
  const L = f.layoutType; const N = f.layoutName;
  const out = (c, confidence, ...reasons) => ({ class: c, confidence, reasons });
  const pc = (v) => `${Math.round(v * 100)}%`;
  if ((f.charts || f.tables) && (f.chartArea >= 0.2 || f.words <= 60 || ['tbl', 'chart', 'txAndChart', 'chartAndTx'].includes(L)))
    return out('chart-table', f.chartArea >= 0.2 ? 'high' : 'medium', f.charts ? `${f.charts} chart object(s)` : `${f.tables} table(s)`, `covers ${pc(f.chartArea)} of the slide`);
  if (i === 0 && (f.hasCtrTitle || L === 'title' || (f.words >= 1 && f.words <= 30 && f.textBlocks <= 3 && f.imageCoverage < 0.5)))
    return out('title', f.hasCtrTitle || L === 'title' ? 'high' : 'medium', f.hasCtrTitle || L === 'title' ? 'title layout / centred-title placeholder' : 'first slide with little text', 'first slide');
  if (i >= n - 2 && i > 0 && f.closingKw && f.words <= 40) return out('closing', 'high', 'closing wording found (thanks, questions, contact)', i === n - 1 ? 'last slide' : 'second-to-last slide');
  if (i === n - 1 && n > 2 && f.words <= 15 && (L === 'title' || f.hasCtrTitle || L === 'secHead')) return out('closing', 'medium', 'last slide using a title-style layout with little text');
  if (i > 0 && (L === 'secHead' || /section|divider|chapter/i.test(N))) return out('section', L === 'secHead' ? 'high' : 'medium', L === 'secHead' ? 'section-header layout' : `layout name "${N}"`);
  if (i > 0 && i < n - 1 && f.hasTitle && f.words <= 10 && f.textBlocks <= 1 && f.images === 0 && !f.bgImage)
    return out('section', 'medium', `title with only ${f.words} words and no pictures`);
  if (f.columns >= 2) {
    const strong = f.compareKw || L === 'twoTxTwoObj' || /compar/i.test(N);
    return out('comparison', strong ? 'high' : 'medium', `${f.columns} side-by-side text blocks`, strong ? 'comparison wording or layout' : 'two-column text without comparison wording');
  }
  if (f.imageCoverage >= 0.3 || (f.bgImage && f.words <= 40)) return out('image', f.imageCoverage >= 0.5 || f.bgImage ? 'high' : 'medium', f.bgImage ? 'full-slide background image' : `pictures cover ${pc(f.imageCoverage)} of the slide`);
  if (f.images >= 1 && ((f.imageCoverage >= 0.15 && f.words <= 60) || IMG_TYPES.includes(L))) return out('image', 'medium', `picture-led layout, ${f.words} words`);
  if (f.words >= 1) return out('text', f.hasBody ? 'high' : 'medium', `${f.words} words`, f.hasBody ? 'body placeholder' : 'free text boxes');
  return out('other', 'low', 'no readable text, pictures, tables or charts');
}

// ---------- main ----------
/** Analyse a parsed deck (output of parsePptx, which keeps the JSZip handle). Returns the Design System JSON. */
export async function extractDesign(parsed) {
  const { zip, slides, width: W, height: H } = parsed;
  const presDoc = await readXml(zip, 'ppt/presentation.xml');
  const presRels = await readRels(zip, 'ppt/presentation.xml');
  const env = { layouts: new Map(), masters: new Map(), W, H, defaultTextStyle: desc(presDoc, 'defaultTextStyle')[0] || null,
    stats: { unappliedMods: new Set(), clrOverride: false, inferredTitles: 0 } };
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
    const cx = master ? master.cx : { theme: EMPTY_THEME, clrMap: DEFAULT_CLRMAP };
    const e = { ...env, cx, rels, layoutDoc: layout?.doc || null, master };
    const own = readBg(doc, cx, rels, env);
    const bg = own ? { ...own, source: 'slide' } : layout?.bg ? { ...layout.bg, source: 'layout' } : master?.bg ? { ...master.bg, source: 'master' } : { type: 'none', source: 'default' };
    const shapes = [];
    walkTree(desc(doc, 'spTree')[0], (b) => b, e, shapes);
    if (inferTitle(shapes, H)) env.stats.inferredTitles++;
    const f = features(shapes, bg, layout, W, H);
    const cls = classify(f, i, slides.length);
    const effBg = bgHex(bg) || (bg.type === 'none' && cx.theme.colors[cx.clrMap.bg1]) || null;
    S.push({ i, ps, shapes, bg, effBg, f, cls, cx, layout });
  }
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
  if (env.stats.clrOverride) limitations.push('Some layouts override the colour mapping; those overrides are not applied.');
  if (env.stats.inferredTitles) limitations.push(`${env.stats.inferredTitles} slide(s) have no title placeholder; the top, largest text box was treated as the title.`);
  limitations.push('Group rotation and flips are ignored when positioning shapes.');

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
    const nonFooter = s.shapes.filter((x) => !FOOTER.has(x.phRaw));
    const content = nonFooter.filter((x) => x.box && ['text', 'image', 'table', 'chart', 'object'].includes(x.kind) && !(x.box.w >= 0.9 * W && x.box.h >= 0.9 * H));
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
        const bgc = sh.fill?.type === 'solid' && sh.fill.hex ? sh.fill.hex : s.effBg;
        if (!bgc) { unevalBg.n += r.chars; }
        else if (r.hex) {
          const ratio = contrast(r.hex, bgc); const large = r.sz >= 18 || (r.b && r.sz >= 14);
          if (ratio < (large ? 3 : 4.5) && (!worst || ratio < worst.ratio)) worst = { ratio: r2(ratio), text: r.hex, background: bgc };
        }
      }
    }
    bgTypes[s.bg.type] = (bgTypes[s.bg.type] || 0) + 1;
    if (s.bg.hex) { addPal(s.bg, 'bg', 1); bump(bgColors, s.bg.hex, 1); } else if (s.effBg && s.bg.type === 'none') { addPal({ hex: s.effBg, slot: 'bg1' }, 'bg', 1); bump(bgColors, s.effBg, 1); }

    // margins from content bounds
    let bounds = null;
    if (content.length) {
      const x0 = Math.min(...content.map((c) => c.box.x)); const y0 = Math.min(...content.map((c) => c.box.y));
      const x1 = Math.max(...content.map((c) => c.box.x + c.box.w)); const y1 = Math.max(...content.map((c) => c.box.y + c.box.h));
      bounds = { left: r1((x0 / W) * 100), top: r1((y0 / H) * 100), right: r1(((W - x1) / W) * 100), bottom: r1(((H - y1) / H) * 100) };
    }
    // vertical gaps between stacked content
    const gaps = []; const stack = content.filter((c) => c.kind !== 'object').sort((a, b) => a.box.y - b.box.y);
    for (let k = 1; k < stack.length; k++) {
      const a = stack[k - 1].box; const b = stack[k].box;
      const ov = Math.min(a.x + a.w, b.x + b.w) - Math.max(a.x, b.x);
      if (ov >= 0.5 * Math.min(a.w, b.w) && b.y >= a.y + a.h) gaps.push(((b.y - a.y - a.h) / H) * 100);
    }
    Object.assign(s, { bounds, minPt, maxPt, titleSz, smallChars, worst, gaps, shapeCount, content });
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
  if (unevalBg.n && totText) limitations.push(`Contrast was not evaluated for ${Math.round((unevalBg.n / totText) * 100)}% of text (gradient, image or effect backgrounds have no single colour).`);
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
    textInsetsEmu: Object.fromEntries(Object.entries(insets).map(([k, v]) => [k, v.length ? median(v) : { l: 91440, t: 45720, r: 91440, b: 45720 }[k]])),
    insetsSource: insets.l.length || insets.t.length ? 'explicit where set, PowerPoint defaults elsewhere' : 'PowerPoint defaults',
    alignment: Object.fromEntries(Object.entries(algn).map(([k, v]) => [k, r2(v / totAlign)])),
  };

  // patterns and recurring elements
  const bin = (b) => `${Math.round((b.x / W) * 20)},${Math.round((b.y / H) * 20)},${Math.round((b.w / W) * 20)},${Math.round((b.h / H) * 20)}`;
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
    slideRecs.push({ n: s.i + 1, title: s.ps.title, layout: s.layout?.name || null, layoutType: s.layout?.type || null, class: s.cls.class, confidence: s.cls.confidence, reasons: s.cls.reasons,
      background: { type: s.bg.type, hex: s.bg.hex || s.effBg || null, source: s.bg.source }, words: s.f.words, counts: s.shapeCount, imageCoverage: r2(s.f.imageCoverage), bounds: s.bounds, minPt: s.minPt, maxPt: s.maxPt });
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
  const add = (code, severity, message, slides = []) => findings.push({ code, severity, message, slides: slides.slice(0, 20) });
  const small = S.filter((s) => s.smallChars > 0).map((s) => s.i + 1);
  if (small.length) add('small-text', 'warn', 'Text under 12 pt is hard to read on a projector.', small);
  if (t.sizes.length && body != null && body < 14) add('low-body-size', 'info', `The dominant body size is ${body} pt; 16 pt or more reads better on screen.`);
  const manyFonts = t.fonts.filter((f) => f.share >= 0.02);
  if (manyFonts.length > 2) add('many-fonts', 'warn', `${manyFonts.length} font families carry more than 2% of the text each (${manyFonts.map((f) => f.name).join(', ')}). Two is the usual limit.`);
  const off = palette.filter((c) => !c.slot && !c.neutral);
  if (off.length >= 3) add('off-theme-colors', 'info', `${off.length} hard-coded colours are not part of the theme palette, so they will not follow a theme change.`);
  const low = S.filter((s) => s.worst).map((s) => s.i + 1);
  if (low.length) add('low-contrast', 'warn', 'Text/background contrast is below WCAG AA (4.5:1 for normal text, 3:1 for large text).', low);
  const lefts = S.filter((s) => s.bounds).map((s) => s.bounds.left); const medL = median(lefts);
  const offMargin = S.filter((s) => s.bounds && Math.abs(s.bounds.left - medL) > 4).map((s) => s.i + 1);
  if (offMargin.length >= 2 && offMargin.length < n) add('inconsistent-margins', 'info', 'Left margins differ by more than 4% of the slide width from the template median.', offMargin);
  const tsz = S.filter((s) => s.cls.class === 'text' && s.titleSz != null); const medT = median(tsz.map((s) => s.titleSz));
  const oddT = tsz.filter((s) => Math.abs(s.titleSz - medT) > 4).map((s) => s.i + 1);
  if (oddT.length) add('inconsistent-title-size', 'info', `Titles on text slides vary by more than 4 pt from the typical ${medT} pt.`, oddT);
  const dense = S.filter((s) => s.f.words > 110).map((s) => s.i + 1);
  if (dense.length) add('dense-text', 'info', 'More than 110 words on one slide.', dense);

  return {
    preserve: { description: 'Keep the template exactly as analysed. Tokens are the observed fonts, sizes, colours and margins; nothing is normalised.', tokens: observed },
    enhance: { description: 'Keep the template\'s fonts, palette and layout identity, but raise sizes, margins and contrast to readable minimums. Only values listed under "changes" differ from Preserve.', tokens: enh, changes, findings },
    note: 'Modes are a stored contract for slide generation (Phase 3). Phase 2 never modifies your PPTX file.',
  };
}
