// Native slide renderer (Phase 3.1). Draws a resolved scene as live DOM: every shape is an absolutely positioned <div> holding an <svg> (geometry, fill, outline,
// picture) and a text <div> (real, selectable, editable-style text); tables are real <table>s. Nothing is rasterised: no <canvas>, no screenshot of the slide.
// Read-only: it never writes to the OOXML. Sizes are expressed in points through one CSS variable (--u = CSS pixels per point), so resizing the slide
// only changes that variable. The CSP forbids inline style attributes, so all styling goes through the CSSOM (el.style.*) and SVG presentation attributes.
// Every element carries data-id (the stable element id), data-kind and data-support (native | approximated | opaque).
import { kid, kids, mediaUrl } from '../pptx/parser.js';
import { note, phSources } from '../model/scene.js';
import { resolveScene } from '../model/resolve.js';
import { presetGeometry, customGeometry } from './geometry.js';
import { fontStack, loadStyles, loadFontsFor, fontReport } from './fonts.js';

const SVG = 'http://www.w3.org/2000/svg';
const E = 12700;
const f = (n) => (Math.round(n * 1000) / 1000).toString();
const L = (pt) => `calc(var(--u) * ${f(pt)})`;
const rgba = (hex, a = 1) => { if (a >= 1) return `#${hex}`; const [r, g, b] = [0, 2, 4].map((i) => parseInt(hex.slice(i, i + 2), 16)); return `rgba(${r},${g},${b},${f(a)})`; };
const el = (tag, cls, attrs) => { const n = document.createElement(tag); if (cls) n.className = cls; for (const [k, v] of Object.entries(attrs || {})) n.setAttribute(k, v); return n; };
const sv = (tag, attrs) => { const n = document.createElementNS(SVG, tag); for (const [k, v] of Object.entries(attrs || {})) n.setAttribute(k, String(v)); return n; };
const DASH = { solid: [], dot: [1, 3], sysDot: [1, 1], dash: [4, 3], sysDash: [3, 1], lgDash: [8, 3], dashDot: [4, 3, 1, 3], sysDashDot: [3, 1, 1, 1], lgDashDot: [8, 3, 1, 3], lgDashDotDot: [8, 3, 1, 3, 1, 3], sysDashDotDot: [3, 1, 1, 1, 1, 1] };
const OPAQUE_LABEL = { chart: 'Chart', diagram: 'SmartArt diagram', ole: 'Embedded object', frame: 'Content', content: 'Ink / external content', unknown: 'Unsupported element' };

let uid = 0;
const mix = (a, b) => [0, 2, 4].map((i) => Math.round((parseInt(a.slice(i, i + 2), 16) + parseInt(b.slice(i, i + 2), 16)) / 2).toString(16).padStart(2, '0')).join('').toUpperCase();

// ---------- paint (SVG fills, strokes, markers) ----------
function gradientDef(defs, fill, w, h) {
  const id = `sfg${++uid}`;
  let g;
  if (fill.kind === 'path') g = sv('radialGradient', { id, cx: '50%', cy: '50%', r: '50%' });
  else {
    const a = (fill.angle * Math.PI) / 180; const dx = Math.cos(a); const dy = Math.sin(a); const len = Math.abs(w * dx) + Math.abs(h * dy); const cx = w / 2; const cy = h / 2;
    g = sv('linearGradient', { id, gradientUnits: 'userSpaceOnUse', x1: f(cx - (dx * len) / 2), y1: f(cy - (dy * len) / 2), x2: f(cx + (dx * len) / 2), y2: f(cy + (dy * len) / 2) });
  }
  for (const s of fill.stops) g.append(sv('stop', { offset: f(s.pos / 100), 'stop-color': `#${s.hex}`, 'stop-opacity': f(s.alpha ?? 1) }));
  defs.append(g);
  return `url(#${id})`;
}
function markerDef(defs, end, color, alpha) {
  const id = `sfm${++uid}`; const size = { sm: 2, med: 3, lg: 5 }; const mw = size[end.w] || 3; const mh = size[end.len] || 3;
  const m = sv('marker', { id, viewBox: '0 0 10 10', refX: end.type === 'arrow' ? 8 : 5, refY: 5, markerWidth: mw, markerHeight: mh, markerUnits: 'strokeWidth', orient: 'auto-start-reverse', overflow: 'visible' });
  const paint = { fill: `#${color}`, 'fill-opacity': f(alpha) };
  if (end.type === 'oval') m.append(sv('circle', { cx: 5, cy: 5, r: 5, ...paint }));
  else if (end.type === 'diamond') m.append(sv('path', { d: 'M0 5L5 0L10 5L5 10Z', ...paint }));
  else if (end.type === 'stealth') m.append(sv('path', { d: 'M0 0L10 5L0 10L3 5Z', ...paint }));
  else if (end.type === 'arrow') m.append(sv('path', { d: 'M1 0L9 5L1 10', fill: 'none', stroke: `#${color}`, 'stroke-width': 1.5, 'stroke-opacity': f(alpha) }));
  else m.append(sv('path', { d: 'M0 0L10 5L0 10Z', ...paint }));
  defs.append(m);
  return `url(#${id})`;
}
function applyStroke(path, ln, defs, open) {
  if (!ln || ln.type !== 'solid') { path.setAttribute('stroke', 'none'); return; }
  const w = Math.max((ln.w ?? 9525) / E, 0.25);
  path.setAttribute('stroke', `#${ln.hex}`); path.setAttribute('stroke-width', f(w));
  if ((ln.alpha ?? 1) < 1) path.setAttribute('stroke-opacity', f(ln.alpha));
  if (ln.cap) path.setAttribute('stroke-linecap', ln.cap);
  path.setAttribute('stroke-linejoin', ln.join || 'round');
  if (ln.join === 'miter') path.setAttribute('stroke-miterlimit', 8);
  const dash = DASH[ln.dash] || [];
  if (dash.length) path.setAttribute('stroke-dasharray', dash.map((d) => f(d * w)).join(' '));
  if (open) { if (ln.head) path.setAttribute('marker-start', markerDef(defs, ln.head, ln.hex, ln.alpha ?? 1)); if (ln.tail) path.setAttribute('marker-end', markerDef(defs, ln.tail, ln.hex, ln.alpha ?? 1)); }
}

// ---------- geometry lookup (own, else the placeholder's on layout / master) ----------
function geometryOf(scene, e, w, h) {
  const spPr = kid(e.el, 'spPr');
  let src = spPr;
  if (!kid(spPr, 'prstGeom') && !kid(spPr, 'custGeom') && e.ph) {
    const { layout, master } = phSources(e.ph, scene.layout?.doc, scene.master?.doc, e.origin);
    for (const s of [layout, master]) { const p = s && kid(s, 'spPr'); if (p && (kid(p, 'prstGeom') || kid(p, 'custGeom'))) { src = p; break; } }
  }
  const cg = kid(src, 'custGeom');
  let r = null;
  if (cg) { r = customGeometry(cg, w, h); if (!r) note(e, 'approximated', 'custom geometry has no drawable path: drawn as its bounding rectangle'); }
  else {
    const pg = kid(src, 'prstGeom'); const prst = pg?.getAttribute('prst') || 'rect';
    const adj = {}; for (const g of kids(kid(pg, 'avLst'), 'gd')) { const m = /^val\s+(-?\d+)/.exec(g.getAttribute('fmla') || ''); if (m) adj[g.getAttribute('name')] = Number(m[1]); }
    r = presetGeometry(prst, w, h, adj);
    if (!r) note(e, 'approximated', `preset shape "${prst}" is not implemented: drawn as its bounding rectangle`);
  }
  if (!r) r = presetGeometry('rect', w, h, {});
  else if (r.approx) note(e, 'approximated', r.approx);
  return r;
}

// ---------- pictures / image fills ----------
async function imageHref(ctx, target) {
  if (!target) return null;
  if (!ctx.urls.has(target)) ctx.urls.set(target, mediaUrl(ctx.zip, target));
  return ctx.urls.get(target);
}
function placeholderPicture(svg, w, h, label) {
  svg.append(sv('rect', { x: 0, y: 0, width: f(w), height: f(h), fill: '#E7E6E6', stroke: '#A6A6A6', 'stroke-dasharray': '4 3', 'stroke-width': 0.75 }));
  svg.append(sv('path', { d: `M0 0L${f(w)} ${f(h)}M${f(w)} 0L0 ${f(h)}`, stroke: '#BFBFBF', 'stroke-width': 0.75, fill: 'none' }));
  if (w > 60 && h > 14) { const t = sv('text', { x: f(w / 2), y: f(h / 2), 'text-anchor': 'middle', 'dominant-baseline': 'middle', fill: '#595959', 'font-size': Math.min(11, w / 12), 'font-family': 'sans-serif' }); t.textContent = label; svg.append(t); }
}
async function paintImage(ctx, e, svg, defs, fill, g, w, h) {
  const href = fill.target && !fill.external && !fill.missing ? await imageHref(ctx, fill.target) : null;
  if (!href) { placeholderPicture(svg, w, h, fill.external ? 'Linked picture (not loaded)' : fill.missing ? 'Picture missing' : `Picture (${(fill.target || '').split('.').pop().toUpperCase() || 'unknown'}) not shown`); return; }
  const cid = `sfc${++uid}`; const cp = sv('clipPath', { id: cid }); cp.append(...g.paths.map((p) => sv('path', { d: p.d })));
  defs.append(cp);
  const c = fill.crop || { l: 0, t: 0, r: 0, b: 0 }; const vw = 1 - (c.l + c.r) / 100000; const vh = 1 - (c.t + c.b) / 100000;
  const iw = w / (vw || 1); const ih = h / (vh || 1);
  const grp = sv('g', { 'clip-path': `url(#${cid})` });
  const img = sv('image', { href, x: f((-c.l / 100000) * iw), y: f((-c.t / 100000) * ih), width: f(iw), height: f(ih), preserveAspectRatio: 'none' });
  if (fill.alpha != null && fill.alpha < 1) img.setAttribute('opacity', f(fill.alpha));
  if (fill.gray) { const id = `sff${++uid}`; const flt = sv('filter', { id }); flt.append(sv('feColorMatrix', { type: 'saturate', values: 0 })); defs.append(flt); img.setAttribute('filter', `url(#${id})`); }
  grp.append(img); svg.append(grp);
}

// ---------- shapes ----------
async function renderShape(ctx, e, box) {
  const { scene } = ctx; const w = e.world.w / E; const h = e.world.h / E;
  const g = geometryOf(scene, e, w, h);
  const sw = Math.max(w, 1); const sh = Math.max(h, 1);
  const svg = sv('svg', { class: 'sf-geo', viewBox: `0 0 ${f(sw)} ${f(sh)}`, preserveAspectRatio: 'none' });
  svg.style.width = L(sw); svg.style.height = L(sh);
  const defs = sv('defs'); svg.append(defs);
  const st = e.style || { fill: { type: 'none' }, line: { type: 'none' }, shadow: null };
  const fill = st.fill;
  if (fill.type === 'image') await paintImage(ctx, e, svg, defs, fill, g, w, h);
  for (const p of g.paths) {
    const path = sv('path', { d: p.d });
    if (p.evenodd) path.setAttribute('fill-rule', 'evenodd');
    if (!p.fill || fill.type === 'none' || fill.type === 'image') path.setAttribute('fill', 'none');
    else if (fill.type === 'solid') { path.setAttribute('fill', `#${fill.hex}`); if ((fill.alpha ?? 1) < 1) path.setAttribute('fill-opacity', f(fill.alpha)); }
    else if (fill.type === 'gradient') path.setAttribute('fill', gradientDef(defs, fill, w, h));
    else if (fill.type === 'pattern') path.setAttribute('fill', `#${mix(fill.fg, fill.bg)}`);
    if (p.stroke) applyStroke(path, st.line, defs, g.open); else path.setAttribute('stroke', 'none');
    svg.append(path);
  }
  if (e.world.flipX) svg.style.transform = 'scaleX(-1)';
  const fx = [];
  if (st.shadow) { const s = st.shadow; fx.push(`drop-shadow(${L(s.dx / E)} ${L(s.dy / E)} ${L(s.blur / E / 2)} ${rgba(s.hex, s.alpha)})`); }
  if (st.glow) { const g = st.glow; fx.push(`drop-shadow(0 0 ${L(g.rad / E / 2)} ${rgba(g.hex, g.alpha)})`, `drop-shadow(0 0 ${L(g.rad / E / 4)} ${rgba(g.hex, g.alpha)})`); }
  if (fx.length) svg.style.filter = fx.join(' ');
  box.append(svg);
  return g;
}

// ---------- text ----------
const ALIGN = { left: 'left', right: 'right', center: 'center', justify: 'justify' };
function paragraphEl(p, k) {
  const d = el('div', 'sf-p'); d.dir = p.rtl ? 'rtl' : 'ltr';
  const st = d.style;
  st.textAlign = ALIGN[p.algn] || 'left';
  if (p.marL) st.paddingInlineStart = L(p.marL / E);
  if (p.indent) st.textIndent = L(p.indent / E);
  const first = p.runs.find((r) => !r.br);
  const size = (first?.sz ?? p.endSz ?? 18);
  // The paragraph's own font is the line's strut: without this the browser's 16 px default makes every line taller than PowerPoint's.
  st.fontSize = `calc(var(--u) * ${f(size)}${k.fit ? ' * var(--fit, 1)' : ''})`;
  st.fontFamily = fontStack(first?.font, first?.rtl);
  if (p.lnSpc?.pct != null) st.lineHeight = f(p.lnSpc.pct * 1.2); else if (p.lnSpc?.pts != null) st.lineHeight = L(p.lnSpc.pts);
  const sp = (s) => (s ? (s.pts != null ? L(s.pts) : L(s.pct * size * 1.2)) : null);
  if (sp(p.spcBef)) st.marginTop = sp(p.spcBef);
  if (sp(p.spcAft)) st.marginBottom = sp(p.spcAft);
  const fit = k.fit ? ' * var(--fit, 1)' : '';
  if (p.bullet) {
    const b = el('span', 'sf-bu'); b.setAttribute('aria-hidden', 'true');
    b.textContent = p.bullet.type === 'auto' ? p.bullet.label : p.bullet.char;
    const bs = p.bullet.sizePts ?? (p.bullet.sizePct ? size * p.bullet.sizePct : size);
    b.style.fontSize = `calc(var(--u) * ${f(bs)}${fit})`;
    if (p.bullet.font) b.style.fontFamily = fontStack(p.bullet.font);
    else if (first) b.style.fontFamily = fontStack(first.font, first.rtl);
    b.style.color = p.bullet.color ? `#${p.bullet.color}` : first ? rgba(first.hex, first.alpha) : '';
    if (p.indent < 0) b.style.width = L(-p.indent / E); else b.style.marginInlineEnd = L(size * 0.3);
    d.append(b);
  }
  for (const r of p.runs) {
    if (r.br) { d.append(document.createElement('br')); continue; }
    const s = el('span', r.link ? 'sf-r sf-link' : 'sf-r'); s.textContent = r.text;
    if (r.rtl) s.lang = 'ar';
    const c = s.style; const rs = (r.sz ?? 18) * (r.baseline ? 0.66 : 1);
    c.fontFamily = fontStack(r.font, r.rtl);
    c.fontSize = `calc(var(--u) * ${f(rs)}${fit})`;
    c.color = rgba(r.hex || '000000', r.alpha ?? 1);
    if (r.b) c.fontWeight = '700';
    if (r.i) c.fontStyle = 'italic';
    const deco = [r.u ? 'underline' : null, r.strike ? 'line-through' : null].filter(Boolean).join(' ');
    if (deco) c.textDecoration = deco;
    if (r.baseline) { c.verticalAlign = `calc(var(--u) * ${f(((r.sz ?? 18) * r.baseline) / 100)}${fit})`; }
    if (r.caps === 'all') c.textTransform = 'uppercase'; else if (r.caps === 'small') c.fontVariantCaps = 'small-caps';
    if (r.spc) c.letterSpacing = L(r.spc);
    if (r.highlight) c.backgroundColor = `#${r.highlight}`;
    d.append(s);
  }
  if (!p.runs.some((r) => !r.br && r.text) && !p.runs.some((r) => r.br)) { const z = el('span', 'sf-r'); z.textContent = '​'; z.style.fontSize = `calc(var(--u) * ${f(size)}${fit})`; d.append(z); }
  return d;
}
function textBox(text, k) {
  const t = text.body; const box = el('div', 'sf-text');
  const s = box.style;
  s.paddingLeft = L(t.insets.l / E); s.paddingRight = L(t.insets.r / E); s.paddingTop = L(t.insets.t / E); s.paddingBottom = L(t.insets.b / E);
  s.justifyContent = { top: 'flex-start', middle: 'center', bottom: 'flex-end' }[t.anchor] || 'flex-start';
  if (t.wrap === 'none') s.whiteSpace = 'nowrap';
  if (t.vert) { s.writingMode = 'vertical-rl'; if (t.vert === 'vert270') s.transform = 'rotate(180deg)'; }
  const inner = el('div', 'sf-text-in');
  if (t.cols > 1) { inner.style.columnCount = String(t.cols); inner.style.columnGap = L(t.colGap / E); inner.style.columnFill = 'auto'; inner.style.height = '100%'; } // the first column fills to the frame height, then the text flows on, as in PowerPoint
  if (t.wrap === 'none') box.dataset.wrap = 'none';
  for (const p of text.paras) inner.append(paragraphEl(p, k));
  box.append(inner);
  if (t.autofit === 'shrink' && t.fontScale === 1) box.dataset.autofit = 'shrink';
  return box;
}

// ---------- tables ----------
function cssBorder(b) {
  if (!b || b.type !== 'solid') return 'none';
  const w = Math.max((b.w ?? 12700) / E, 0.5);
  return `${L(w)} ${b.dash && /dot|Dot/.test(b.dash) ? 'dotted' : b.dash && b.dash !== 'solid' ? 'dashed' : 'solid'} ${rgba(b.hex, b.alpha ?? 1)}`;
}
function renderTable(ctx, e, box) {
  const T = e.table; if (!T) return;
  const tbl = el('table', 'sf-table'); if (T.rtl) tbl.dir = 'rtl';
  const total = T.cols.reduce((a, b) => a + b, 0) / E;
  tbl.style.width = L(total);
  const cg = document.createElement('colgroup');
  for (const c of T.cols) { const col = document.createElement('col'); col.style.width = L(c / E); cg.append(col); }
  tbl.append(cg);
  for (const row of T.rows) {
    const tr = document.createElement('tr'); tr.style.height = L(row.h / E);
    for (const c of row.cells) {
      if (c.hMerge || c.vMerge) continue;
      const td = document.createElement('td');
      if (c.gridSpan > 1) td.colSpan = c.gridSpan; if (c.rowSpan > 1) td.rowSpan = c.rowSpan;
      const s = td.style;
      if (c.fill?.type === 'solid') s.backgroundColor = rgba(c.fill.hex, c.fill.alpha ?? 1);
      else if (c.fill?.type === 'gradient') s.backgroundColor = rgba(c.fill.stops[0].hex, 1);
      for (const [side, name] of [['left', 'Left'], ['right', 'Right'], ['top', 'Top'], ['bottom', 'Bottom']]) s[`border${name}`] = cssBorder(c.bdr[side] || c.bdr[side === 'left' || side === 'right' ? 'insideV' : 'insideH']);
      s.padding = `${L(c.mar.t / E)} ${L(c.mar.r / E)} ${L(c.mar.b / E)} ${L(c.mar.l / E)}`;
      s.verticalAlign = { top: 'top', middle: 'middle', bottom: 'bottom' }[c.anchor] || 'top';
      for (const p of c.text) td.append(paragraphEl(p, ctx.k));
      tr.append(td);
    }
    tbl.append(tr);
  }
  box.append(tbl);
}

// ---------- one element ----------
const rendersAsShape = (e) => ['shape', 'image', 'line'].includes(e.kind);
async function renderElement(ctx, e) {
  const { scene } = ctx; const W = e.world;
  const box = el('div', 'sf-el'); box.dataset.id = e.id; box.dataset.kind = e.kind; box.dataset.support = e.support; box.dataset.z = String(e.z);
  if (e.name) box.dataset.name = e.name;
  const x = W.cx / E - W.w / E / 2; const y = W.cy / E - W.h / E / 2;
  const s = box.style;
  s.left = L(x); s.top = L(y); s.width = L(W.w / E); s.height = L(W.h / E);
  if (W.rot) s.transform = `rotate(${f(W.rot)}deg)`;
  if (Math.abs(W.shear || 0) > 1e-3) note(e, 'approximated', 'the shape is skewed in the file: drawn without the skew');
  if (e.hidden) { box.hidden = true; box.dataset.hidden = '1'; }
  if (e.descr) { box.setAttribute('aria-label', e.descr); if (e.kind === 'image') box.setAttribute('role', 'img'); }
  ctx.nodes.set(e.id, box);

  if (e.tag === 'grpSp') { box.classList.add('sf-group'); return box; }
  if (rendersAsShape(e)) {
    const g = await renderShape(ctx, e, box);
    if (e.text && e.text.paras.length && (e.text.paras.some((p) => p.runs.length) || false)) {
      const tb = textBox(e.text, ctx.k);
      const r = g.textRect; if (r) { const t = tb.style; t.left = L(r[0]); t.top = L(r[1]); t.width = L(r[2] - r[0]); t.height = L(r[3] - r[1]); }
      box.append(tb);
    }
  } else if (e.kind === 'table') renderTable(ctx, e, box);
  else {
    box.classList.add('sf-opaque');
    const lab = el('div', 'sf-opaque-label'); lab.textContent = `${OPAQUE_LABEL[e.kind] || 'Element'} — not drawn`; const sub = el('div', 'sf-opaque-sub'); sub.textContent = e.name || ''; lab.append(sub);
    box.append(lab);
  }
  box.dataset.support = e.support; // notes added while drawing
  return box;
}

// ---------- background ----------
function paintBackground(ctx, root, bg) {
  const s = root.style;
  if (bg.type === 'solid') s.backgroundColor = rgba(bg.hex, bg.alpha ?? 1);
  else if (bg.type === 'gradient') {
    const stops = bg.stops.map((t) => `${rgba(t.hex, t.alpha ?? 1)} ${f(t.pos)}%`).join(', ');
    s.backgroundImage = bg.kind === 'path' ? `radial-gradient(circle at center, ${stops})` : `linear-gradient(${f(bg.angle + 90)}deg, ${stops})`;
    s.backgroundColor = rgba(bg.stops[0].hex, 1);
  } else if (bg.type === 'pattern') s.backgroundColor = `#${mix(bg.fg, bg.bg)}`;
  else if (bg.type === 'image') return imageHref(ctx, bg.target).then((u) => { if (u && !bg.external) { s.backgroundImage = `url("${u}")`; s.backgroundSize = '100% 100%'; } else s.backgroundColor = '#E7E6E6'; });
  return null;
}

/**
 * Render one scene into `host`. Returns { root, nodes, report, setWidth(px), fitText(), dispose() }.
 * `host` must be in the document (text is measured for "shrink on overflow"). The scene is resolved first when it has not been.
 */
export async function renderSlide(scene, host, opts = {}) {
  await resolveScene(scene);
  await loadStyles();
  await loadFontsFor(scene);
  const ctx = { scene, zip: scene.parsed.zip, urls: new Map(), nodes: new Map(), k: { fit: true } };
  const root = el('div', 'sf-slide'); root.dataset.slide = String(scene.index);
  const wpt = scene.width / E; const hpt = scene.height / E;
  root.style.width = L(wpt); root.style.height = L(hpt);
  root.dir = 'ltr';
  await paintBackground(ctx, root, scene.backgroundFill);
  const bgSupport = scene.backgroundFill.type === 'pattern' ? 'approximated' : scene.backgroundFill.type === 'gradient' && scene.backgroundFill.kind === 'path' ? 'approximated' : scene.backgroundFill.type === 'image' && (scene.backgroundFill.external || scene.backgroundFill.missing || scene.backgroundFill.tile) ? 'approximated' : 'native';
  const layers = {};
  for (const e of scene.all) {
    if (!e.rendered || !e.world || e.kind === 'alt') continue;
    const node = await renderElement(ctx, e);
    root.append(node);
    layers[e.origin] = (layers[e.origin] || 0) + 1;
  }
  host.replaceChildren(root);
  const handle = {
    root, nodes: ctx.nodes,
    setWidth(px) { if (px > 0) { root.style.setProperty('--u', `${f(px / wpt)}px`); handle.fitText(); } },
    fitText() {
      for (const box of root.querySelectorAll('.sf-text[data-autofit="shrink"]')) {
        const owner = box.parentElement; const e = scene.byId.get(owner.dataset.id); const inner = box.firstElementChild;
        box.style.removeProperty('--fit'); let fit = 1;
        const over = () => inner.scrollHeight - 1 > box.clientHeight - parseFloat(getComputedStyle(box).paddingTop) - parseFloat(getComputedStyle(box).paddingBottom) + 0.5;
        if (box.clientHeight > 0 && over()) {
          while (fit > 0.3 && over()) { fit = Math.round((fit - 0.05) * 100) / 100; box.style.setProperty('--fit', String(fit)); }
          if (e) note(e, 'approximated', `text overflows its frame and the frame is set to "shrink on overflow": the text is scaled to ${Math.round(fit * 100)}% here, PowerPoint re-fits it when the file is opened`);
          owner.dataset.support = e?.support || owner.dataset.support;
        }
      }
    },
    dispose() { for (const p of ctx.urls.values()) p.then((u) => u && URL.revokeObjectURL(u)); ctx.urls.clear(); },
  };
  handle.report = () => ({ ...renderReport(scene, bgSupport, layers), fonts: fontReport(scene) });
  if (opts.width) handle.setWidth(opts.width);
  return handle;
}

/** Support summary of everything on the slide (elements of every layer, undrawn alternatives and unplaced elements included). */
export function renderReport(scene, bgSupport = 'native', layers = {}) {
  const counts = { native: 0, approximated: 0, opaque: 0 }; const byKind = {}; const issues = [];
  for (const e of scene.all) {
    counts[e.support]++;
    (byKind[e.kind] ||= { native: 0, approximated: 0, opaque: 0 })[e.support]++;
    if (e.support !== 'native') issues.push({ id: e.id, kind: e.kind, name: e.name, support: e.support, reasons: e.reasons.map((r) => r.text) });
  }
  return { counts, byKind, issues, background: bgSupport, unplaced: scene.unplaced.map((e) => e.id), layers, total: scene.all.length };
}
