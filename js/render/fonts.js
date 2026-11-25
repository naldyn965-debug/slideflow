// Font handling for the Rendered tab (Phase 3.1).
// The deck's own font names are always tried first, so a machine that has Calibri / Arial / Cambria shows them. When it does not, the stack falls back to a
// self-hosted, metric-compatible, properly licensed substitute (see /fonts/README.md and css/fonts.css), then to a generic family. Arabic text always has an
// Arabic-capable font in its stack. Nothing is ever fetched from a third-party host (the CSP only allows 'self'). The self-hosted families carry an "SF " prefix.

/** PowerPoint font (lower-case) -> self-hosted substitute with the same metrics, so line breaks match. */
export const SUBSTITUTES = {
  calibri: 'SF Carlito', 'calibri light': 'SF Carlito', cambria: 'SF Caladea', 'cambria math': 'SF Caladea',
  arial: 'SF Liberation Sans', 'arial narrow': 'SF Liberation Sans', helvetica: 'SF Liberation Sans', 'helvetica neue': 'SF Liberation Sans', 'segoe ui': 'SF Liberation Sans', tahoma: 'SF Liberation Sans', verdana: 'SF Liberation Sans',
  'times new roman': 'SF Liberation Serif', times: 'SF Liberation Serif', georgia: 'SF Liberation Serif', 'courier new': 'SF Liberation Mono', consolas: 'SF Liberation Mono',
};
const GENERIC = { serif: ['cambria', 'times new roman', 'times', 'georgia', 'garamond', 'palatino', 'book antiqua', 'constantia', 'caladea', 'liberation serif', 'traditional arabic', 'simplified arabic'], mono: ['courier new', 'consolas', 'courier', 'lucida console', 'liberation mono'] };
/** Arabic-capable families, in order: nicer system fonts when installed, then the self-hosted face that is always there. */
export const ARABIC_FALLBACK = ['Noto Naskh Arabic', 'Noto Sans Arabic', 'SF DejaVu Sans'];
export const ARABIC_SELF_HOSTED = 'SF DejaVu Sans';
export const STYLESHEETS = ['css/fonts.css', 'css/render.css'];

const generic = (name) => { const n = (name || '').toLowerCase(); return GENERIC.mono.includes(n) ? 'monospace' : GENERIC.serif.includes(n) ? 'serif' : 'sans-serif'; };
const q = (n) => `"${String(n).replace(/["\\]/g, '')}"`;

/** CSS font-family value for a run. `font` is the typeface the deck asked for (may be null), `rtl` marks an Arabic / Hebrew segment. */
export function fontStack(font, rtl = false) {
  const out = [];
  if (font && !/^\+(mj|mn)-/.test(font)) { out.push(q(font)); const s = SUBSTITUTES[font.toLowerCase()]; if (s) out.push(q(s)); }
  else out.push(q('SF Carlito'));
  if (rtl) for (const a of ARABIC_FALLBACK) out.push(q(a));
  out.push(generic(font));
  return out.join(', ');
}

/** Add the font and slide stylesheets once (same-origin <link>s, allowed by the CSP). Resolves when both are parsed. */
export function loadStyles(base = new URL('../../', import.meta.url).href) {
  if (typeof document === 'undefined') return Promise.resolve();
  return Promise.all(STYLESHEETS.map((href) => {
    const id = `sf-css-${href.replace(/\W+/g, '-')}`;
    const have = document.getElementById(id);
    if (have) return have.sheet ? Promise.resolve() : new Promise((r) => { have.addEventListener('load', r, { once: true }); have.addEventListener('error', r, { once: true }); });
    return new Promise((res) => { const l = document.createElement('link'); l.id = id; l.rel = 'stylesheet'; l.href = base + href; l.addEventListener('load', res, { once: true }); l.addEventListener('error', res, { once: true }); document.head.append(l); });
  }));
}

// document.fonts.check() is true for any family unknown to the page, so we compare real text widths against generic references instead.
let canvas;
const localCache = new Map();
export function isInstalled(name) {
  if (typeof document === 'undefined') return false;
  const key = name.toLowerCase();
  if (localCache.has(key)) return localCache.get(key);
  canvas ||= document.createElement('canvas').getContext('2d');
  const text = 'mmmmmmmmmmlli0Oiوعربي';
  const w = (family) => { canvas.font = `40px ${family}`; return canvas.measureText(text).width; };
  const v = ['monospace', 'serif', 'sans-serif'].some((g) => w(`${q(name)}, ${g}`) !== w(g));
  localCache.set(key, v);
  return v;
}

/** Every (font, bold, italic, rtl) combination used by the text of a scene. */
export function fontsUsed(scene) {
  const seen = new Map();
  const add = (r) => { if (r.br || !r.text) return; const k = `${r.font}|${!!r.b}|${!!r.i}|${!!r.rtl}`; if (!seen.has(k)) seen.set(k, { font: r.font, b: !!r.b, i: !!r.i, rtl: !!r.rtl }); };
  for (const e of scene.all) {
    for (const p of e.text?.paras || []) { p.runs.forEach(add); if (p.defRun) add({ ...p.defRun, text: 'x' }); } // (the default run of an empty paragraph: its font is needed once text is typed into it)
    for (const row of e.table?.rows || []) for (const c of row.cells) for (const p of c.text) p.runs.forEach(add);
  }
  return [...seen.values()];
}

/** Which of the deck's fonts are installed, which use a self-hosted substitute, and which fall back to a generic family. */
export function fontReport(scene) {
  const rows = new Map();
  for (const u of fontsUsed(scene)) {
    const name = u.font || '(theme default)';
    const r = rows.get(name) || { name, original: false, substitute: null, arabic: false, generic: false };
    if (u.font && isInstalled(u.font)) r.original = true;
    else if (SUBSTITUTES[(u.font || 'calibri').toLowerCase()]) r.substitute = SUBSTITUTES[(u.font || 'calibri').toLowerCase()];
    else r.generic = true;
    if (u.rtl) r.arabic = true;
    rows.set(name, r);
  }
  return [...rows.values()];
}

/** Download the self-hosted faces this scene needs (only for fonts that are not installed) before text is measured. Resolves after `ms` at the latest; never throws. */
export async function loadFontsFor(scene, ms = 4000) {
  if (typeof document === 'undefined' || !document.fonts) return;
  const jobs = new Set();
  for (const u of fontsUsed(scene)) {
    const spec = (fam, sample) => `${u.i ? 'italic ' : ''}${u.b ? 700 : 400} 16px ${q(fam)}|${sample}`;
    const sub = SUBSTITUTES[(u.font || 'calibri').toLowerCase()];
    if (sub && !(u.font && isInstalled(u.font))) jobs.add(spec(sub, u.rtl ? 'a' : 'abcXYZ019'));
    if (!u.font) jobs.add(spec('SF Carlito', 'abcXYZ019'));
    if (u.rtl) jobs.add(`${u.b ? 700 : 400} 16px ${q(ARABIC_SELF_HOSTED)}|\u0627\u0644\u0639\u0631\u0628\u064a\u0629`);
  }
  const all = Promise.all([...jobs].map((j) => { const [font, text] = j.split('|'); return document.fonts.load(font, text).catch(() => []); }));
  try { await Promise.race([all, new Promise((r) => setTimeout(r, ms))]); await document.fonts.ready; } catch { /* ignore */ }
}
