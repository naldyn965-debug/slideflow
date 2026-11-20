// Scene model (Phase 3.1): a READ-ONLY, in-memory description of one slide, derived from the PPTX XML and never stored anywhere.
// The PPTX XML stays the single source of truth: every element keeps a live reference (`el`) to its XML node, a stable id, its document-order
// z-index and its place in the group tree. Nothing here mutates the DOM, and nothing is dropped: whatever cannot be drawn is still an element,
// marked `opaque` (or `approximated`) with the reason, and elements that have no position are listed in `scene.unplaced`.
//
// Geometry: every element gets `world`, the affine matrix from its own box (0..w, 0..h, origin top-left) to slide EMU, composed through any
// number of nested groups (chOff/chExt scale, flips, rotation, in PowerPoint's order: flip about the centre, then rotate clockwise).
//
// Support levels (set here for structure, extended by js/model/resolve.js and js/render/render.js for features):
//   native       every attribute that affects how it looks is understood and drawn
//   approximated drawn, but something is simplified (the reasons say what)
//   opaque       not drawn from its content: a placeholder box stands in; the XML is untouched and the element can still be located
import { kid, kids, desc, readXml, readRels, placeholderOf, normType } from '../pptx/parser.js';
import { loadMaster, loadLayout, phShape, clrMapOf, hidesMasterShapes, DEFAULT_CLRMAP, EMPTY_THEME } from '../pptx/design.js';

const NS_R = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
const NS_GD = { table: 'http://schemas.openxmlformats.org/drawingml/2006/table', chart: 'http://schemas.openxmlformats.org/drawingml/2006/chart', diagram: 'http://schemas.openxmlformats.org/drawingml/2006/diagram', ole: 'http://schemas.openxmlformats.org/presentationml/2006/ole' };
const LEVEL = { native: 0, approximated: 1, opaque: 2 };
const NOT_SHAPES = new Set(['nvGrpSpPr', 'grpSpPr', 'extLst']);
const SHAPE_TAGS = new Set(['sp', 'pic', 'cxnSp', 'graphicFrame', 'grpSp', 'contentPart']);
export const SUPPORT_LEVELS = ['native', 'approximated', 'opaque'];

// ---------- affine matrices [a, b, c, d, e, f]: x' = a x + c y + e, y' = b x + d y + f ----------
export const IDENTITY = [1, 0, 0, 1, 0, 0];
export const mul = (m, n) => [m[0] * n[0] + m[2] * n[1], m[1] * n[0] + m[3] * n[1], m[0] * n[2] + m[2] * n[3], m[1] * n[2] + m[3] * n[3], m[0] * n[4] + m[2] * n[5] + m[4], m[1] * n[4] + m[3] * n[5] + m[5]];
export const apply = (m, x, y) => [m[0] * x + m[2] * y + m[4], m[1] * x + m[3] * y + m[5]];
const T = (x, y) => [1, 0, 0, 1, x, y];
const S = (x, y) => [x, 0, 0, y, 0, 0];
const R = (deg) => { const t = (deg * Math.PI) / 180; const c = Math.cos(t); const s = Math.sin(t); return [c, s, -s, c, 0, 0]; };

/** Own transform: box-local (0..w, 0..h) -> parent space. Flip about the centre first, then rotate clockwise about it. */
export function localMatrix(x) {
  const cx = x.x + x.w / 2; const cy = x.y + x.h / 2;
  return mul(mul(mul(T(cx, cy), R(x.rot || 0)), S(x.flipH ? -1 : 1, x.flipV ? -1 : 1)), T(-x.w / 2, -x.h / 2));
}
/** A group's child space -> its parent space: map chOff/chExt onto the group box, then flip and rotate the whole group about its centre. */
export function groupMatrix(x) {
  const c = x.child || { x: x.x, y: x.y, w: x.w, h: x.h };
  const sx = c.w ? x.w / c.w : 1; const sy = c.h ? x.h / c.h : 1;
  const U = mul(mul(T(x.x, x.y), S(sx, sy)), T(-c.x, -c.y));
  const gcx = x.x + x.w / 2; const gcy = x.y + x.h / 2;
  const Rg = mul(mul(mul(T(gcx, gcy), R(x.rot || 0)), S(x.flipH ? -1 : 1, x.flipV ? -1 : 1)), T(-gcx, -gcy));
  return mul(Rg, U);
}

/** Split a world matrix into what a renderer needs: centre, final frame size, rotation, mirroring and whether it is sheared (non-uniform scale under a rotation). */
export function decompose(m, w, h) {
  const [cx, cy] = apply(m, w / 2, h / 2);
  const sx = Math.hypot(m[0], m[1]); const sy = Math.hypot(m[2], m[3]);
  const det = m[0] * m[3] - m[1] * m[2];
  const f = det < 0 ? -1 : 1;
  const rot = ((Math.atan2(f * m[1], f * m[0]) * 180) / Math.PI + 360) % 360;
  const shear = sx && sy ? Math.abs(m[0] * m[2] + m[1] * m[3]) / (sx * sy) > 1e-3 : false;
  return { cx, cy, w: w * sx, h: h * sy, rot: Math.abs(rot - 360) < 1e-9 ? 0 : rot, flipX: f < 0, shear };
}

// ---------- reading transforms ----------
function readXfrmEl(x) {
  if (!x) return null;
  const off = kid(x, 'off'); const ext = kid(x, 'ext');
  if (!off || !ext) return null;
  const n = (e, a) => Number(e.getAttribute(a)) || 0;
  const out = { x: n(off, 'x'), y: n(off, 'y'), w: n(ext, 'cx'), h: n(ext, 'cy'), rot: (Number(x.getAttribute('rot')) || 0) / 60000, flipH: /^(1|true)$/.test(x.getAttribute('flipH') || ''), flipV: /^(1|true)$/.test(x.getAttribute('flipV') || '') };
  const co = kid(x, 'chOff'); const ce = kid(x, 'chExt');
  if (co && ce) out.child = { x: n(co, 'x'), y: n(co, 'y'), w: n(ce, 'cx'), h: n(ce, 'cy') };
  return out;
}
/** The element's own <a:xfrm> / <p:xfrm> (graphicFrame, contentPart) or null. */
export function ownXfrm(el) {
  const n = el.localName;
  if (n === 'graphicFrame' || n === 'contentPart') return readXfrmEl(kid(el, 'xfrm'));
  const holder = n === 'grpSp' ? kid(el, 'grpSpPr') : kid(el, 'spPr');
  return readXfrmEl(holder && kid(holder, 'xfrm'));
}

// ---------- context (loaded once per deck, cached) ----------
const CTX = new WeakMap();
/** Deck-level context: slide docs (one DOM per part, kept), theme / master / layout via the Phase 2 loaders. */
export async function deckContext(parsed) {
  let c = CTX.get(parsed);
  if (c) return c;
  const { zip } = parsed;
  const presDoc = await readXml(zip, 'ppt/presentation.xml');
  const env = { layouts: new Map(), masters: new Map(), W: parsed.width, H: parsed.height, defaultTextStyle: desc(presDoc, 'defaultTextStyle')[0] || null, layerCache: new Map(), stats: { unappliedMods: new Set(), inferredTitles: 0 } };
  c = { zip, parsed, presDoc, env, docs: new Map(), rels: new Map(), firstMaster: null };
  CTX.set(parsed, c);
  return c;
}
async function partDoc(ctx, part) {
  if (!ctx.docs.has(part)) ctx.docs.set(part, await readXml(ctx.zip, part));
  return ctx.docs.get(part);
}
async function partRels(ctx, part) {
  if (!ctx.rels.has(part)) ctx.rels.set(part, await readRels(ctx.zip, part));
  return ctx.rels.get(part);
}
async function firstMaster(ctx) {
  if (ctx.firstMaster) return ctx.firstMaster;
  const presRels = await partRels(ctx, 'ppt/presentation.xml');
  const r = [...presRels.values()].find((x) => x.kind === 'slideMaster');
  ctx.firstMaster = r ? await loadMaster(ctx.zip, r.target, ctx.env) : null;
  return ctx.firstMaster;
}

// ---------- support bookkeeping ----------
/** Record why an element is not fully native. Escalates only (native < approximated < opaque). */
export function note(e, level, text) {
  if (LEVEL[level] > LEVEL[e.support]) e.support = level;
  if (text && !e.reasons.some((r) => r.text === text)) e.reasons.push({ level, text });
}

const nvOf = (el) => [...el.children].find((c) => /^nv/.test(c.localName)) || null;
function identOf(el) {
  const nv = nvOf(el);
  const c = nv ? kid(nv, 'cNvPr') : desc(el, 'cNvPr')[0]; // elements this model does not know may carry a cNvPr anywhere
  return c ? { xmlId: c.getAttribute('id'), name: c.getAttribute('name') || '', descr: c.getAttribute('descr') || '', hidden: /^(1|true)$/.test(c.getAttribute('hidden') || '') } : { xmlId: null, name: '', descr: '', hidden: false };
}

function frameKind(el) {
  const gd = desc(el, 'graphicData')[0];
  const uri = gd?.getAttribute('uri') || '';
  for (const [k, v] of Object.entries(NS_GD)) if (uri === v) return k;
  if (desc(el, 'tbl').length) return 'table';
  if (desc(el, 'oleObj').length) return 'ole';
  return 'frame';
}

// ---------- the walker ----------
function makeElement(ctx, st, el, path, parent) {
  const n = el.localName;
  const idn = identOf(el);
  const e = {
    id: null, part: st.part, origin: st.origin, path, depth: parent ? parent.depth + 1 : 0, parentId: parent ? parent.id : null, childIds: [],
    tag: n, kind: 'unknown', name: idn.name, xmlId: idn.xmlId, descr: idn.descr, hidden: idn.hidden, el, ph: null, xfrm: null, posSource: null, world: null,
    support: 'native', reasons: [], rendered: st.rendered, branch: st.branch, z: st.z.n++, hasText: false, geom: null, image: null,
  };
  // ---- id: part + cNvPr id; duplicates and missing ids are made unique and flagged, never merged ----
  let id = idn.xmlId != null && idn.xmlId !== '' ? `${st.part}#${idn.xmlId}` : null;
  if (!id) { id = `${st.part}@${path.join('/')}`; e.idIssue = 'missing'; }
  if (st.branch && st.branch.suffix) id += st.branch.suffix;
  if (st.ids.has(id)) { let k = 2; while (st.ids.has(`${id}~${k}`)) k++; id = `${id}~${k}`; e.idIssue = e.idIssue || 'duplicate'; }
  st.ids.add(id); e.id = id;
  if (e.idIssue) note(e, 'approximated', e.idIssue === 'missing' ? 'element has no id in the file (a path-based id is used)' : 'duplicate id in the file (made unique here)');
  if (!e.rendered) note(e, 'opaque', st.branch?.reason || 'not drawn');

  // ---- kind ----
  const ph = placeholderOf(el);
  e.ph = ph;
  if (n === 'sp') { e.kind = 'shape'; e.geom = kid(kid(el, 'spPr'), 'prstGeom')?.getAttribute('prst') || (kid(kid(el, 'spPr'), 'custGeom') ? 'custom' : null); e.hasText = desc(el, 't').some((t) => t.textContent !== ''); }
  else if (n === 'pic') { e.kind = 'image'; e.geom = kid(kid(el, 'spPr'), 'prstGeom')?.getAttribute('prst') || null; }
  else if (n === 'cxnSp') { e.kind = 'line'; e.geom = kid(kid(el, 'spPr'), 'prstGeom')?.getAttribute('prst') || (kid(kid(el, 'spPr'), 'custGeom') ? 'custom' : null); }
  else if (n === 'grpSp') e.kind = 'group';
  else if (n === 'graphicFrame') e.kind = frameKind(el);
  else if (n === 'contentPart') e.kind = 'content';
  else e.kind = 'unknown';

  // ---- media references (pictures and picture fills) ----
  const blip = n === 'pic' || n === 'sp' ? desc(el, 'blip')[0] : null;
  if (blip) {
    const rid = blip.getAttributeNS(NS_R, 'embed') || blip.getAttributeNS(NS_R, 'link');
    const rel = rid && st.rels.get(rid);
    e.image = rel ? { rid, target: rel.target, external: rel.external } : { rid: rid || null, target: null, external: false, missing: true };
  }

  // ---- position: own xfrm, else the layout / master placeholder it inherits from ----
  let xf = ownXfrm(el); e.posSource = xf ? 'own' : null;
  if (!xf && !SHAPE_TAGS.has(n)) { const x = desc(el, 'xfrm')[0]; xf = readXfrmEl(x); e.posSource = xf ? 'own' : null; }
  if (!xf && ph) {
    const srcs = phSources(ph, st.layoutDoc, st.masterDoc, st.origin);
    for (const [src, shape] of [['layout', srcs.layout], ['master', srcs.master]]) {
      const x = shape && ownXfrm(shape);
      if (x) { xf = x; e.posSource = src; break; }
    }
  }
  e.xfrm = xf;
  return e;
}

/** The layout and master shapes a placeholder inherits from. Slide -> layout matches by idx (then type); layout -> master matches by TYPE only, because idx is
 *  only meaningful between a slide and its layout (a master's date placeholder is idx 2, and so is a layout's second content placeholder). The Phase 2 helper
 *  phShape() matches by idx on both hops; it is left as is so the design analysis output does not change. */
const MASTER_TYPE = { ctrTitle: 'title', subTitle: 'body', obj: 'body', chart: 'body', tbl: 'body', pic: 'body', media: 'body', clipArt: 'body', dgm: 'body', body: 'body' };
export function phSources(ph, layoutDoc, masterDoc, origin = 'slide') {
  if (!ph) return { layout: null, master: null };
  const layout = origin === 'slide' ? phShape(ph, layoutDoc) : null;
  const lp = layout && placeholderOf(layout);
  const t = lp?.type || ph.type; const want = MASTER_TYPE[t] || normType(t);
  let master = null;
  if (origin !== 'master' && masterDoc) for (const sp of desc(masterDoc, 'sp')) { const p = placeholderOf(sp); if (p && (MASTER_TYPE[p.type] || normType(p.type)) === want) { master = sp; break; } }
  return { layout, master };
}

function place(e, M) {
  const xf = e.xfrm;
  if (!xf) return false;
  const W = mul(M, localMatrix(xf));
  const corners = [[0, 0], [xf.w, 0], [xf.w, xf.h], [0, xf.h]].flatMap(([u, v]) => apply(W, u, v));
  e.world = { matrix: W, corners, ...decompose(W, xf.w, xf.h) };
  if (e.world.shear) note(e, 'approximated', 'non-uniform group scale under a rotation skews this element in PowerPoint; drawn without the skew');
  return true;
}

function walk(ctx, st, tree, parent, M, out) {
  const list = [...tree.children];
  for (let i = 0; i < list.length; i++) {
    const el = list[i]; const n = el.localName;
    if (NOT_SHAPES.has(n)) continue;
    const path = [...(parent ? parent.path : st.basePath), i];
    if (n === 'AlternateContent') { walkAlt(ctx, st, el, path, parent, M, out); continue; }
    const e = makeElement(ctx, st, el, path, parent);
    out.push(e);
    if (parent) parent.childIds.push(e.id);
    st.byId.set(e.id, e);
    classify(e);
    if (n === 'grpSp') {
      let childM = M;
      if (e.xfrm) { place(e, M); childM = mul(M, groupMatrix(e.xfrm)); } else note(e, 'approximated', 'group has no transform; its children are drawn in the parent coordinate space');
      walk(ctx, st, el, e, childM, out);
      continue;
    }
    if (!place(e, M)) { note(e, 'opaque', e.ph ? 'placeholder with no position in the slide, layout or master' : 'no position in the file'); st.unplaced.push(e); }
  }
}

/** mc:AlternateContent: recorded as a wrapper element. Branch choice follows what a consumer that understands plain DrawingML shapes does (PowerPoint 2010+ for
 *  p14 / p15 content): the first Choice made only of ordinary shapes, pictures, connectors and groups is drawn; a Choice that needs a feature we cannot draw
 *  (a chart-ex or other graphic frame, unrecognised elements) hands over to the Fallback; with no Fallback the first Choice is drawn.
 *  Every branch's elements are kept in the model (rendered: false for those not drawn), so nothing is lost. */
const PLAIN = new Set(['sp', 'pic', 'cxnSp', 'grpSp']);
const plainChoice = (c) => { const k = [...c.children]; return k.length > 0 && k.every((x) => PLAIN.has(x.localName)); };
function walkAlt(ctx, st, el, path, parent, M, out) {
  const choices = kids(el, 'Choice'); const fb = kid(el, 'Fallback');
  const wrapper = {
    id: `${st.part}@${path.join('/')}`, part: st.part, origin: st.origin, path, depth: parent ? parent.depth + 1 : 0, parentId: parent ? parent.id : null, childIds: [], tag: 'AlternateContent', kind: 'alt', name: 'AlternateContent', xmlId: null, descr: '', hidden: false, el, ph: null, xfrm: null, posSource: null, world: null,
    support: 'native', reasons: [], rendered: st.rendered, branch: st.branch, z: st.z.n++, hasText: false, geom: null, image: null,
  };
  const pick = choices.find(plainChoice);
  const draw = pick || fb || choices[0];
  const why = pick ? 'drawn from the Choice branch (ordinary shapes)' : fb ? 'the Choice needs a feature that is not drawn here: drawn from the Fallback branch' : 'no Fallback: the first Choice is drawn';
  note(wrapper, 'approximated', `mc:AlternateContent (Requires="${choices.map((c) => c.getAttribute('Requires')).filter(Boolean).join(', ')}"): ${why}; the other branch is kept untouched in the file`);
  out.push(wrapper); st.byId.set(wrapper.id, wrapper); if (parent) parent.childIds.push(wrapper.id);
  const branches = [...choices.map((c, k) => ({ node: c, name: `choice${k ? k + 1 : ''}`, suffix: k ? `~choice${k + 1}` : '' })), ...(fb ? [{ node: fb, name: 'fallback', suffix: '~fallback' }] : [])];
  for (const b of branches) {
    const drawn = b.node === draw;
    const sub = { ...st, rendered: st.rendered && drawn, branch: { name: b.name, suffix: b.suffix, reason: `mc:AlternateContent ${b.name} branch is not drawn (kept untouched in the file)` } };
    const holder = { id: wrapper.id, path: [...path, [...el.children].indexOf(b.node)], depth: wrapper.depth, childIds: wrapper.childIds };
    walk(ctx, sub, b.node, { ...holder }, M, out);
  }
}

/** Support from structure alone. Style- and feature-level downgrades are added later by resolve.js / render.js through note(). */
function classify(e) {
  if (e.kind === 'chart') note(e, 'opaque', 'chart: not drawn from its data (placeholder box); the chart part is untouched');
  else if (e.kind === 'diagram') note(e, 'opaque', 'SmartArt: not drawn; the diagram parts are untouched');
  else if (e.kind === 'ole') note(e, 'opaque', 'embedded object (OLE): not drawn; the object is untouched');
  else if (e.kind === 'frame') note(e, 'opaque', 'unrecognised graphic frame content: not drawn; the XML is untouched');
  else if (e.kind === 'content') note(e, 'opaque', 'content part (ink or other external content): not drawn; the XML is untouched');
  else if (e.kind === 'unknown') note(e, 'opaque', `unrecognised element <${e.tag}>: not drawn; the XML is untouched`);
  if (e.kind === 'image') {
    const nv = nvOf(e.el); const nvPr = nv && kid(nv, 'nvPr');
    if (nvPr && ['videoFile', 'audioFile', 'wavAudioFile', 'quickTimeFile'].some((t) => kid(nvPr, t))) note(e, 'approximated', 'audio/video: only the poster image is drawn');
  }
}

// ---------- layers ----------
function layerElements(ctx, part, doc, origin, rels, layoutDoc, masterDoc, z) {
  const tree = desc(doc, 'spTree')[0];
  const st = { part, origin, rendered: true, branch: null, rels, layoutDoc, masterDoc, ids: new Set(), byId: new Map(), unplaced: [], z, basePath: [] };
  const out = [];
  if (tree) walk(null, st, tree, null, IDENTITY, out);
  // Template placeholders are inheritance sources, not slide content: they are never drawn on a slide, so they are left out of the layer.
  const keep = origin === 'slide' ? out : out.filter((e) => !(e.ph && e.depth === 0));
  return { origin, part, rels, elements: keep, byId: st.byId, unplaced: st.unplaced };
}

/** Build the scene of one slide. `index` is the 0-based position in parsed.slides. */
export async function buildScene(parsed, index) {
  const ctx = await deckContext(parsed);
  const ps = parsed.slides[index];
  if (!ps) throw new RangeError(`No slide at index ${index}`);
  const doc = await partDoc(ctx, ps.part);
  const rels = await partRels(ctx, ps.part);
  const lRel = [...rels.values()].find((r) => r.kind === 'slideLayout');
  const layout = lRel ? await loadLayout(ctx.zip, lRel.target, ctx.env) : null;
  const master = layout?.master || (await firstMaster(ctx));
  if (layout?.doc) ctx.docs.set(layout.path, layout.doc);
  if (master?.doc) ctx.docs.set(master.path, master.doc);
  const slideOvr = clrMapOf(doc);
  const clrMap = { ...(master?.clrMap || DEFAULT_CLRMAP), ...(layout?.clrOvr || {}), ...(slideOvr || {}) };
  const cx = { theme: master ? master.theme : EMPTY_THEME, clrMap };

  const z = { n: 0 };
  const layers = [];
  const hideAll = hidesMasterShapes(doc);
  if (!hideAll) {
    if (master?.doc && !hidesMasterShapes(layout?.doc)) layers.push(layerElements(ctx, master.path, master.doc, 'master', master.rels, null, null, z));
    if (layout?.doc) layers.push(layerElements(ctx, layout.path, layout.doc, 'layout', layout.rels, layout.doc, master?.doc || null, z));
  }
  const slideLayer = layerElements(ctx, ps.part, doc, 'slide', rels, layout?.doc || null, master?.doc || null, z);
  layers.push(slideLayer);

  // background: the first <p:bg> found on slide -> layout -> master
  const bgOf = (d) => kid(desc(d, 'cSld')[0], 'bg');
  let background = { source: 'default', el: null, part: null, rels: null };
  for (const [source, owner] of [['slide', { doc, part: ps.part, rels }], ['layout', layout && { doc: layout.doc, part: layout.path, rels: layout.rels }], ['master', master && { doc: master.doc, part: master.path, rels: master.rels }]]) {
    const bg = owner?.doc && bgOf(owner.doc);
    if (bg) { background = { source, el: bg, part: owner.part, rels: owner.rels }; break; }
  }

  const byId = new Map(); const all = [];
  for (const l of layers) { for (const [k, v] of l.byId) byId.set(k, v); all.push(...l.elements); }
  return {
    parsed, index, part: ps.part, width: parsed.width, height: parsed.height, doc, rels, layoutPart: layout?.path || null, masterPart: master?.path || null,
    cx, layout, master, hidden: ps.hidden, layers, background, elements: slideLayer.elements, all, byId, unplaced: slideLayer.unplaced,
    masterShapesHidden: hideAll || hidesMasterShapes(layout?.doc),
  };
}

/** The relationships of the part an element lives in (slide, layout or master): needed to resolve its pictures. */
export const relsOf = (scene, e) => scene.layers.find((l) => l.part === e.part)?.rels || scene.rels;

// ---------- helpers for consumers and tests ----------
/** Counts of slide elements by support level and kind. Group wrappers are counted like any other element. */
export function sceneStats(scene) {
  const s = { total: 0, native: 0, approximated: 0, opaque: 0, byKind: {}, unplaced: scene.unplaced.length };
  for (const e of scene.elements) { s.total++; s[e.support]++; s.byKind[e.kind] = (s.byKind[e.kind] || 0) + 1; }
  return s;
}
/** Every shape-like element that exists in the XML of a part's spTree, counted independently of the walker (for the "nothing is dropped" check). */
export function countXmlElements(doc) {
  const tree = desc(doc, 'spTree')[0];
  let n = 0;
  const visit = (parent) => {
    for (const c of parent.children) {
      if (NOT_SHAPES.has(c.localName)) continue;
      if (c.localName === 'AlternateContent') { n++; for (const b of c.children) visit(b); continue; }
      n++;
      if (c.localName === 'grpSp') visit(c);
    }
  };
  if (tree) visit(tree);
  return n;
}
