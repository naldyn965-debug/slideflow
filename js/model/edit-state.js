// Temporary edit-state (Phase 3.2.3-B moves, Phase 3.2.4-B resizes, Phase 3.2.5-B text, Phase 3.2.6-A styles): the place where COMPLETED drag movements, COMPLETED resizes, COMPLETED text edits and COMPLETED style edits are kept. It is a small in-memory model that sits NEXT TO the Scene Model and never inside it:
//   - the PPTX XML stays the single source of truth and is never touched (no xfrm is written, nothing is exported, nothing is saved: this model lives and dies with the open preview);
//   - the Scene Model (scene.byId, e.world, e.xfrm, the DOM of the slide parts) is only READ here; a scene is always built from the untouched file, so a scene rebuilt later gives the same
//     ids and the same file positions, and the edit-state is simply applied on top of it again (that is how a move survives a redraw, another slide, or another tab);
//   - only a COMPLETED move is committed (a pointerup of the dragging pointer). A cancelled drag never reaches this module, so an earlier committed move is exactly what it was.
//
// What is stored per moved element (one record per element id, the TOTAL move since the file, not a list of steps):
//   slideDelta   { dx, dy }  the move in SLIDE space, in points (what the preview draws; exact at every zoom because it is scaled by --u, never stored in pixels)
//   localDelta   { dx, dy }  the same move in the element's PARENT coordinate space, in EMU: the slide itself for a top-level element, the child space (chOff / chExt) of the
//                            parent group for a group child, through any number of nested groups. This is the change an element's own <a:off> would need in the file.
//   origin       { x, y }    the element's own <a:off> in the file (EMU, same space as localDelta)
//   position     { x, y }    origin + localDelta
//
// RESIZES (Phase 3.2.4-B) are kept next to the moves, in their own records (one per element id, the TOTAL resize since the file), so a move record is exactly what it was in 3.2.3-B:
//   slideFrame   { w, h, tx, ty }   what the drag preview draws, in slide POINTS: the frame's size in its own (rotated) axes, and the shift of its centre from the file frame's centre (the resize alone: a move is separate)
//   scale        { sx, sy }         slideFrame size / the file frame's world size, per axis (an axis the file gives under FLAT_PT is not scaled: 1)
//   origin       { x, y, w, h }     the element's own <a:off> and <a:ext> in the file (EMU, parent coordinate space)
//   extent       { w, h }           the resized <a:ext> in the same space: origin.w * sx, origin.h * sy
//   localShift   { dx, dy }         the shift of the frame's centre in the PARENT space (EMU): the same inverse of the parent's linear part as for a move (a group's rotation, flips and non-uniform scale are undone)
//   position     { x, y }           the resized <a:off> when the element is not moved: origin.x + localShift.dx - (extent.w - origin.w) / 2 (the centre moves by localShift, the size changes about the centre)
// Why this is exact: the preview stretches the frame about its centre in its own axes. The element's own rotation and flips act about its centre, so they leave centre and axes alone; the parent's
// child space only scales the box: world size = own extent * |column of the world matrix|, which does not depend on the extent. So scaling the own extent by (sx, sy) scales the drawn frame by the
// same factors, and the centre goes where the slide shift says through the same L^-1 as a move. Rotation and flips stay in the file as they are, nothing about the group is touched.
// TEXT edits (Phase 3.2.5-B) are kept in a third map, one record per element id, the TOTAL text change since the file, only the runs that changed:
//   paras   [ { p, runs: [ { r, text, was } ], fill?: { text } } ]
//           p = index into the resolved paragraphs (e.text.paras[p], the same order as the <a:p> of the shape), r = index into e.text.paras[p].runs (br entries count), text = the run's new text,
//           was = its text in the file (a record is only applied where the file text is still `was`), fill = text typed into an EMPTY paragraph (no run to hold it): it becomes one run with the
//           paragraph's default run (para.defRun: its <a:endParaRPr> + the inheritance chain) as its formatting.
// Only text strings are stored. Every run keeps the formatting it resolved to (size, font, colour, link, rtl flag, ...), paragraphs keep their direction, alignment, bullet and level, line breaks stay
// where they are. effectiveText(e) returns the resolved text with the record applied as NEW objects (the Scene Model is never written); a run's `src` ({ ci, s }: the child index of its <a:r> in the <a:p>
// and the script segment of its text; a fill run has src.def) says where a later export would write the text. Nothing here writes XML.
//
// STYLE edits (Phase 3.2.6-A) are kept in a fourth map, one record per element id, the TOTAL style change since the file, only what differs from the file's resolved value. Nothing here draws
// or writes anything: the Scene Model, the renderer, the controller, the UI and the XML are not touched by this module; a later phase reads these records.
//   text   { color?, size?, font?, b?, i? }   applied to EVERY run of the shape (and to the default run of an empty paragraph, so text typed there has it): color = 'RRGGBB' (upper case, alpha 1 as
//                                             PowerPoint writes it; hyperlink runs keep the theme's hyperlink colour), size = NOMINAL points (what <a:rPr sz> would hold, 1..4000, two decimals:
//                                             the resolved run.sz of a shape with a stored normAutofit fontScale is already multiplied by it), font = a typeface name, b / i = booleans
//   fill   { hex }                            a SOLID fill, 'RRGGBB' (replaces any fill the file has: solid, gradient, pattern, none; a picture fill is refused)
//   line   { hex }                            the colour of an outline that exists and is solid (width, dash, caps and joins stay as the file has them)
//   was    what the file has, for what the record holds: text.{color,size,font,b,i} (the value every affected run has in the file, null when the runs differ or the face is the theme's),
//          fill { type, hex, alpha }, line { hex, alpha } - so a later reader can tell a stale record from a current one.
// commitStyle(scene, id, patch) takes { text?, fill?, line? }; inside text: color / size / font / b / i. A value sets, `null` reverts that property (or, as text / fill / line themselves, the whole
// group), an absent key leaves it. A value equal to what the file has (for text properties: what EVERY affected run has) is not stored. No record is left when nothing differs. Everything is checked
// before anything is changed: a refused patch changes nothing. effectiveStyle(e, rec) and effectiveText(e, rec, styleRec) give the element as it looks with the records applied (new objects).
//
// frame(id) puts a move and a resize of one element together: the <a:off> / <a:ext> its own xfrm would hold in the file (the move and the resize are both centre shifts and add up).
//
// Why the local move is exact: an element's world matrix is W = M * Lm, with M = parent child space -> slide and Lm = the element's own box transform (translate to its centre, rotate,
// flip). Moving the box by l in the parent space gives W' = T(L * l) * W where L is the linear part of M (M's translation drops out). So the slide move d is reached by l = L^-1 * d: the
// group's rotation, flips and (non-uniform) scale are undone exactly, and the element's OWN rotation and flips never enter (they act about its centre, which only translates).
// M is taken from the scene as W * Lm^-1, so it holds for every element the scene placed, whatever put it there (own xfrm, inherited placeholder, nested groups, alternate content).
// A parent whose scale is zero (a collapsed group) has no inverse: such a move is refused ('singular') and the element stays where the file has it.
import { localMatrix, mul } from './scene.js';
import { MIN_PT, FLAT_PT, MAX_TIMES, SAME_PT } from './resize-limits.js';
import { paraRtl, isBlankText, defaultFrame, shapeFrame, ADDED_KINDS, KIND_LABEL } from './textbox.js';

const E = 12700;
const num = (v) => typeof v === 'number' && Number.isFinite(v);

/** Linear part [a, b, c, d] (x' = a x + c y, y' = b x + d y) of the parent child space -> slide matrix of a placed element, or null when the scene has not placed it. */
export function parentLinear(e) {
  if (!e?.world?.matrix || !e.xfrm) return null;
  const Lm = localMatrix(e.xfrm); const det = Lm[0] * Lm[3] - Lm[1] * Lm[2];
  if (!num(det) || det === 0) return null;
  // Lm^-1 (linear part), then M = W * Lm^-1 (the translations do not matter for a delta)
  const inv = [Lm[3] / det, -Lm[1] / det, -Lm[2] / det, Lm[0] / det, 0, 0];
  const W = e.world.matrix; const M = mul([W[0], W[1], W[2], W[3], 0, 0], inv);
  return [M[0], M[1], M[2], M[3]];
}

/** A move in slide space (points) -> the same move in the element's parent coordinate space (EMU), or null when that space has no inverse (see the header). */
export function toLocalDelta(e, dxPt, dyPt) {
  const L = parentLinear(e); if (!L) return null;
  const [a, b, c, d] = L; const det = a * d - b * c;
  const n = Math.hypot(a, b) * Math.hypot(c, d);
  if (!num(det) || !(n > 0) || Math.abs(det) <= 1e-12 * n) return null;
  const sx = dxPt * E; const sy = dyPt * E;
  const out = { dx: (d * sx - c * sy) / det, dy: (-b * sx + a * sy) / det };
  return num(out.dx) && num(out.dy) ? out : null;
}

const copy = (r) => ({ ...r, slideDelta: { ...r.slideDelta }, localDelta: { ...r.localDelta }, origin: { ...r.origin }, position: { ...r.position } });
const copyRz = (r) => ({ ...r, slideFrame: { ...r.slideFrame }, scale: { ...r.scale }, origin: { ...r.origin }, extent: { ...r.extent }, localShift: { ...r.localShift }, position: { ...r.position } });
const EPS = 1e-6; // points: float noise of a size that went through the pointer and back
const copyAdded = (r) => ({ ...r, ...(r.paras ? { paras: r.paras.map((p) => ({ ...p })) } : {}) });
const copyTx = (r) => ({ ...r, paras: r.paras.map((p) => ({ ...p, runs: p.runs.map((x) => ({ ...x })), ...(p.fill ? { fill: { ...p.fill } } : {}) })) });
/** The text of an element with a text record applied: new paragraph / run objects, `e.text` itself is never written. Without a record it IS `e.text`. */
function textWithRecord(e, rec) {
  if (!e || !e.text || !rec) return e ? e.text : null;
  return {
    ...e.text,
    paras: e.text.paras.map((p, pi) => {
      const pr = rec.paras.find((x) => x.p === pi); if (!pr) return p;
      let runs = p.runs;
      if (pr.runs.length) runs = p.runs.map((r, ri) => { const c = pr.runs.find((x) => x.r === ri); return c && !r.br && r.text === c.was ? { ...r, text: c.text } : r; });
      if (pr.fill && p.defRun && p.runs.length === 0) runs = [{ ...p.defRun, text: pr.fill.text, fill: true }];
      return runs === p.runs ? p : { ...p, runs };
    }),
  };
}
/**
 * The text of an element with a text record applied (as before 3.2.6-A) and, when `styleRec` (a style record, Phase 3.2.6-A) carries text properties, those applied to every run and to the default
 * run of each paragraph, as NEW objects. `size` is nominal: the run's `sz` becomes size * the stored fontScale, the same scale the resolver put into the file's runs. Without a style record the result is
 * exactly what it was before.
 */
export function effectiveText(e, rec, styleRec) {
  const t = textWithRecord(e, rec);
  const st = styleRec && styleRec.text;
  if (!t || !t.paras || !st) return t;
  const fs = fontScaleOf(t);
  const tr = (r) => {
    if (r.br) return r;
    const o = { ...r };
    if (st.color != null && !r.link) { o.hex = st.color; o.alpha = 1; }
    if (st.size != null) o.sz = st.size * fs;
    if (st.font != null) o.font = st.font;
    if (st.b != null) o.b = st.b;
    if (st.i != null) o.i = st.i;
    return o;
  };
  return { ...t, paras: t.paras.map((p) => ({ ...p, runs: p.runs.map(tr), ...(p.defRun ? { defRun: tr(p.defRun) } : {}) })) };
}

// ---------- style edits (Phase 3.2.6-A) ----------
const HEX = /^#?([0-9a-f]{6})$/i;
const SIZE_MIN = 1; const SIZE_MAX = 4000; // the range of <a:rPr sz> (100..400000 hundredths of a point)
const SIZE_TOL = 0.1; // points: the resolver rounds a run's size to 0.1 after scaling it, so a nominal size recovered from it can be off by a little
const TEXT_KEYS = ['color', 'size', 'font', 'b', 'i'];
const copySt = (r) => ({ ...r, ...(r.text ? { text: { ...r.text } } : {}), ...(r.fill ? { fill: { ...r.fill } } : {}), ...(r.line ? { line: { ...r.line } } : {}), was: { ...(r.was.text ? { text: { ...r.was.text } } : {}), ...(r.was.fill ? { fill: { ...r.was.fill } } : {}), ...(r.was.line ? { line: { ...r.was.line } } : {}) } });
const plain = (v) => v !== null && typeof v === 'object' && !Array.isArray(v);
const fontScaleOf = (t) => { const s = t && t.body && t.body.fontScale; return num(s) && s > 0 ? s : 1; };
const hexOf = (v) => { const m = typeof v === 'string' ? HEX.exec(v) : null; return m ? m[1].toUpperCase() : null; };
const sizeOf = (v) => (num(v) && v >= SIZE_MIN && v <= SIZE_MAX ? Math.round(v * 100) / 100 : null);
const fontOf = (v) => { if (typeof v !== 'string') return null; const t = v.trim(); return t && t.length <= 100 && t[0] !== '+' && !/[\u0000-\u001f\u007f"'\\<>&]/.test(t) ? t : null; };
const alphaOne = (a) => (a == null ? 1 : a) >= 0.9999;
const styleOf = (e) => e.style || { fill: { type: 'none' }, line: { type: 'none' }, shadow: null };

/** The runs a text style edit reaches in the FILE's text: every run, and the default run of a paragraph that has none. */
function textTargets(e) {
  const out = [];
  for (const p of e.text.paras) { let any = false; for (const r of p.runs) if (!r.br) { any = true; out.push(r); } if (!any && p.defRun) out.push(p.defRun); }
  return out;
}
/** What the file has for each text property: { n, color: { n, value }, size, font, b, i } where value is what every affected run has, null when they differ. */
function fileText(e) {
  const rs = textTargets(e); const fs = fontScaleOf(e.text);
  const common = (list, eq) => ({ n: list.length, value: list.length && list.every((v) => eq(v, list[0])) ? list[0] : null });
  const same = (a, b) => a === b;
  return {
    n: rs.length,
    color: common(rs.filter((r) => !r.link).map((r) => (r.hex ? String(r.hex).toUpperCase() : null)), same),
    size: common(rs.map((r) => (num(r.sz) ? Math.round((r.sz / fs) * 100) / 100 : null)), (a, b) => a != null && b != null && Math.abs(a - b) < SIZE_TOL),
    font: common(rs.map((r) => (r.font ? String(r.font) : null)), (a, b) => a != null && b != null && a.toLowerCase() === b.toLowerCase()),
    b: common(rs.map((r) => !!r.b), same), i: common(rs.map((r) => !!r.i), same),
  };
}
/** Merge `patch` into the old record's sections and check every value. Returns { ok: true, text, fill, line, set } or { ok: false, reason: 'invalid' }; nothing is changed. */
function parsePatch(patch, old) {
  const bad = { ok: false, reason: 'invalid' };
  if (!plain(patch)) return bad;
  const keys = Object.keys(patch); if (!keys.length || keys.some((k) => k !== 'text' && k !== 'fill' && k !== 'line')) return bad;
  const out = { ok: true, text: old && old.text ? { ...old.text } : {}, fill: old && old.fill ? { ...old.fill } : null, line: old && old.line ? { ...old.line } : null, set: { text: false, fill: false, line: false } };
  if ('text' in patch && patch.text !== undefined) {
    const t = patch.text;
    if (t === null) out.text = {};
    else {
      if (!plain(t)) return bad; const tk = Object.keys(t); if (!tk.length || tk.some((k) => !TEXT_KEYS.includes(k))) return bad;
      for (const k of tk) {
        const v = t[k];
        if (v === undefined) continue;
        if (v === null) { delete out.text[k]; continue; }
        const x = k === 'color' ? hexOf(v) : k === 'size' ? sizeOf(v) : k === 'font' ? fontOf(v) : typeof v === 'boolean' ? v : null;
        if (x === null) return bad;
        out.text[k] = x; out.set.text = true;
      }
    }
  }
  for (const k of ['fill', 'line']) {
    if (!(k in patch) || patch[k] === undefined) continue;
    const v = patch[k];
    if (v === null) { out[k] = null; continue; }
    if (!plain(v) || Object.keys(v).length !== 1 || !('hex' in v)) return bad;
    const x = hexOf(v.hex); if (x === null) return bad;
    out[k] = { hex: x }; out.set[k] = true;
  }
  return out;
}

/** The style of an element with a style record applied: the resolved fill / line with the edited colour (new objects; `e.style` itself is never written). Without a record it IS `e.style`. */
export function effectiveStyle(e, rec) {
  if (!e) return null;
  if (!rec || (!rec.fill && !rec.line)) return e.style;
  const st = styleOf(e);
  return {
    ...st,
    fill: rec.fill ? { type: 'solid', hex: rec.fill.hex, alpha: 1 } : st.fill,
    line: rec.line ? { ...st.line, type: 'solid', hex: rec.line.hex, alpha: 1 } : st.line,
    shadow: st.shadow ? { ...st.shadow } : st.shadow,
    ...(st.glow ? { glow: { ...st.glow } } : {}),
  };
}

/** A new, empty edit-state. */
export function createEditState() {
  const moves = new Map(); // element id -> record
  const resizes = new Map(); // element id -> resize record (Phase 3.2.4-B)
  const texts = new Map(); // element id -> text record (Phase 3.2.5-B)
  const styles = new Map(); // element id -> style record (Phase 3.2.6-A)
  const added = new Map(); // element id -> record of a new text box (Phase 3.2.7-A) or shape (3.2.7-B): { id, part, xmlId, kind, name, x, y, w, h (EMU), paras: [{ text, rtl }] (text boxes only) }; the element's own content, see js/model/textbox.js
  let version = 0; // counts every change of the content, so "nothing was written" is checkable

  return {
    /**
     * Keep the COMPLETED move of one element: `dx` / `dy` is its total offset from the file position in slide points (the preview offset). Replaces an earlier record of the same
     * element; a zero offset removes it (the element is back at its file position). Returns { ok: true, edit, changed } (edit is null for the zero offset) or { ok: false, reason }.
     * Nothing is changed when it is refused.
     */
    commit(scene, id, dx, dy) {
      const e = scene?.byId?.get(id);
      if (!e || e.origin !== 'slide' || e.kind === 'group' || e.kind === 'alt') return { ok: false, reason: 'not-movable' };
      if (!num(dx) || !num(dy)) return { ok: false, reason: 'invalid' };
      if (!dx && !dy) { const had = moves.delete(id); if (had) version++; return { ok: true, edit: null, changed: had }; }
      const local = toLocalDelta(e, dx, dy);
      if (!local) return { ok: false, reason: e.world && e.xfrm ? 'singular' : 'unplaced' };
      const origin = { x: e.xfrm.x, y: e.xfrm.y };
      const rec = {
        id, part: e.part, slideIndex: scene.index, parentId: e.parentId || null, kind: e.kind, posSource: e.posSource,
        slideDelta: { dx, dy }, localDelta: local, origin, position: { x: origin.x + local.dx, y: origin.y + local.dy },
      };
      const old = moves.get(id);
      const same = old && old.slideDelta.dx === dx && old.slideDelta.dy === dy && old.origin.x === origin.x && old.origin.y === origin.y;
      if (!same) { moves.set(id, rec); version++; }
      return { ok: true, edit: copy(moves.get(id)), changed: !same };
    },
    /**
     * Keep the COMPLETED resize of one element (Phase 3.2.4-B): `frame` = { w, h, tx, ty } is the frame the drag preview draws, in slide points (see the header); null (or a frame equal to the file's)
     * removes the record, the element is back at its file size. It replaces an earlier record of the same element. Returns { ok: true, edit, changed } (edit is null for no resize) or
     * { ok: false, reason }, and nothing is changed when it is refused. The size limits are checked here as well as in the drag (MIN_PT / FLAT_PT / MAX_TIMES), so a frame the drag could never have
     * produced is not kept: 'not-resizable' (not a drawn slide element, or a group), 'invalid' (not finite, negative), 'too-small', 'too-large', 'flat-axis' (an axis the file gives flat was changed),
     * 'singular' / 'unplaced' (as for a move: the parent space cannot be inverted / the scene did not place the element).
     */
    commitResize(scene, id, frame) {
      const e = scene?.byId?.get(id);
      if (!e || e.origin !== 'slide' || e.kind === 'group' || e.kind === 'alt') return { ok: false, reason: 'not-resizable' };
      if (frame == null) { const had = resizes.delete(id); if (had) version++; return { ok: true, edit: null, changed: had }; }
      if (!['w', 'h', 'tx', 'ty'].every((k) => num(frame[k])) || frame.w < 0 || frame.h < 0) return { ok: false, reason: 'invalid' }; // a flat axis (a line's zero height) is 0 and valid; a scaled axis is checked against the limits below
      const W = e.world;
      if (!W || !e.xfrm) return { ok: false, reason: 'unplaced' };
      const fw = W.w / E; const fh = W.h / E; const flatW = !(fw >= FLAT_PT); const flatH = !(fh >= FLAT_PT);
      const sameAsFile = (Math.abs(frame.w - fw) < SAME_PT && Math.abs(frame.h - fh) < SAME_PT && Math.abs(frame.tx) < SAME_PT && Math.abs(frame.ty) < SAME_PT);
      if (sameAsFile) { const had = resizes.delete(id); if (had) version++; return { ok: true, edit: null, changed: had }; }
      if (!num(fw) || !num(fh)) return { ok: false, reason: 'unplaced' };
      if ((flatW && Math.abs(frame.w - fw) > EPS) || (flatH && Math.abs(frame.h - fh) > EPS)) return { ok: false, reason: 'flat-axis' };
      if ((!flatW && frame.w < Math.min(MIN_PT, fw) - EPS) || (!flatH && frame.h < Math.min(MIN_PT, fh) - EPS)) return { ok: false, reason: 'too-small' };
      if (frame.w > Math.max(MAX_TIMES * (scene.width / E), fw) + EPS || frame.h > Math.max(MAX_TIMES * (scene.height / E), fh) + EPS) return { ok: false, reason: 'too-large' };
      const sx = flatW ? 1 : frame.w / fw; const sy = flatH ? 1 : frame.h / fh;
      const shift = toLocalDelta(e, frame.tx, frame.ty);
      if (!shift) return { ok: false, reason: 'singular' };
      const origin = { x: e.xfrm.x, y: e.xfrm.y, w: e.xfrm.w, h: e.xfrm.h };
      const extent = { w: origin.w * sx, h: origin.h * sy };
      if (!num(extent.w) || !num(extent.h) || extent.w < 0 || extent.h < 0) return { ok: false, reason: 'invalid' };
      const rec = {
        id, part: e.part, slideIndex: scene.index, parentId: e.parentId || null, kind: e.kind, posSource: e.posSource,
        slideFrame: { w: frame.w, h: frame.h, tx: frame.tx, ty: frame.ty }, scale: { sx, sy }, origin, extent, localShift: shift,
        position: { x: origin.x + shift.dx - (extent.w - origin.w) / 2, y: origin.y + shift.dy - (extent.h - origin.h) / 2 },
      };
      const old = resizes.get(id); const o = old && old.slideFrame;
      const same = old && o.w === frame.w && o.h === frame.h && o.tx === frame.tx && o.ty === frame.ty && old.origin.x === origin.x && old.origin.y === origin.y && old.origin.w === origin.w && old.origin.h === origin.h;
      if (!same) { resizes.set(id, rec); version++; }
      return { ok: true, edit: copyRz(resizes.get(id)), changed: !same };
    },
    /** A copy of the resize record of an element, or null. */
    getResize(id) { const r = resizes.get(id); return r ? copyRz(r) : null; },
    hasResize(id) { return resizes.has(id); },
    /**
     * Keep the COMPLETED text edit of one element (Phase 3.2.5-B): `draft` = { paras: [ { p, runs: [ { r, text } ], fill? } ] } is the TOTAL text change since the file (the changed runs only; `fill` = text for an empty
     * paragraph). Entries that equal the file's text are dropped; a draft with nothing left removes the record (the element is back at the file's text). Returns { ok: true, edit, changed } (edit is null for
     * no change) or { ok: false, reason } with nothing changed: 'not-editable' (not a slide text shape), 'invalid' (an index, a type or a run that is not there, a `fill` for a paragraph that has runs or no default run).
     */
    commitText(scene, id, draft) {
      const e = scene?.byId?.get(id);
      if (!e || e.origin !== 'slide' || e.tag !== 'sp' || e.kind !== 'shape' || !e.text || !e.text.paras) return { ok: false, reason: 'not-editable' };
      if (!draft || !Array.isArray(draft.paras)) return { ok: false, reason: 'invalid' };
      const out = [];
      for (const dp of draft.paras) {
        const p = e.text.paras[dp && dp.p]; if (!p || !Number.isInteger(dp.p) || !Array.isArray(dp.runs)) return { ok: false, reason: 'invalid' };
        const runs = []; const seen = new Set();
        for (const dr of dp.runs) {
          const r = p.runs[dr && dr.r]; if (!r || r.br || !Number.isInteger(dr.r) || typeof dr.text !== 'string' || seen.has(dr.r)) return { ok: false, reason: 'invalid' };
          seen.add(dr.r); if (dr.text !== r.text) runs.push({ r: dr.r, text: dr.text, was: r.text });
        }
        let fill = null;
        if (dp.fill != null) {
          const t = typeof dp.fill === 'string' ? dp.fill : dp.fill.text;
          if (typeof t !== 'string' || p.runs.length !== 0 || !p.defRun) return { ok: false, reason: 'invalid' };
          if (t !== '') fill = { text: t };
        }
        if (runs.length || fill) out.push({ p: dp.p, runs: runs.sort((a, b) => a.r - b.r), ...(fill ? { fill } : {}) });
      }
      out.sort((a, b) => a.p - b.p);
      if (!out.length) { const had = texts.delete(id); if (had) version++; return { ok: true, edit: null, changed: had }; }
      const rec = { id, part: e.part, slideIndex: scene.index, paras: out };
      const old = texts.get(id); const same = old && JSON.stringify(old.paras) === JSON.stringify(rec.paras);
      if (!same) { texts.set(id, rec); version++; }
      return { ok: true, edit: copyTx(texts.get(id)), changed: !same };
    },
    /** A copy of the text record of an element, or null. */
    getText(id) { const r = texts.get(id); return r ? copyTx(r) : null; },
    hasText(id) { return texts.has(id); },
    /** Copies of the text records of one slide part. */
    textsForSlide(part) { return [...texts.values()].filter((r) => r.part === part).map(copyTx); },
    listTexts() { return [...texts.values()].map(copyTx); },
    /** Forget the text edit of one element (it has the file's text again; a move or a resize of it stays). */
    discardText(id) { if (texts.delete(id)) version++; },
    /** The text of an element with its committed text edit applied (new objects; the scene is not written): see effectiveText. */
    effectiveText(e) { return effectiveText(e, e ? texts.get(e.id) : null); },
    /**
     * Keep the style edit of one element (Phase 3.2.6-A): `patch` = { text?: { color?, size?, font?, b?, i? }, fill?: { hex }, line?: { hex } }, see the header. A value sets, null reverts, a value
     * equal to the file's is not kept, and an element whose record ends up empty has none. Returns { ok: true, edit, changed } (edit = a copy of the record, null when none is left) or
     * { ok: false, reason } with NOTHING changed: 'not-styleable' (not a plain slide shape or connector line), 'invalid' (a malformed patch or value), 'no-text' (a text property for an element
     * with no run to hold it), 'no-color-target' (a text colour for text that is only hyperlinks), 'no-fill-target' (a fill for a connector line), 'picture-fill' (a fill for a shape whose
     * fill is a picture), 'no-outline' (a line colour for an element whose outline is not solid or is not there). Reverting (null) is always allowed for a styleable element.
     */
    commitStyle(scene, id, patch) {
      const e = scene?.byId?.get(id);
      const shape = e && e.origin === 'slide' && e.tag === 'sp' && e.kind === 'shape'; const connector = e && e.origin === 'slide' && e.tag === 'cxnSp' && e.kind === 'line';
      if (!shape && !connector) return { ok: false, reason: 'not-styleable' };
      const old = styles.get(id) || null;
      const p = parsePatch(patch, old); if (!p.ok) return p;
      const st = styleOf(e);
      if (p.set.text) {
        const hasText = shape && e.text && Array.isArray(e.text.paras) && textTargets(e).length > 0;
        if (!hasText) return { ok: false, reason: 'no-text' };
        if (patch.text && patch.text.color != null && fileText(e).color.n === 0) return { ok: false, reason: 'no-color-target' };
      }
      if (p.set.fill) { if (connector) return { ok: false, reason: 'no-fill-target' }; if (st.fill && st.fill.type === 'image') return { ok: false, reason: 'picture-fill' }; }
      if (p.set.line && !(st.line && st.line.type === 'solid')) return { ok: false, reason: 'no-outline' };
      // what the file has, for the properties that stay: a value equal to it is dropped
      const text = {}; const wasText = {}; let ft = null;
      for (const k of TEXT_KEYS) {
        if (!(k in p.text)) continue;
        ft ||= shape && e.text && e.text.paras ? fileText(e) : null;
        const fv = ft ? ft[k].value : null; const v = p.text[k];
        const eq = fv != null && (k === 'size' ? Math.abs(fv - v) < SIZE_TOL : k === 'font' ? String(fv).toLowerCase() === v.toLowerCase() : fv === v);
        if (!eq) { text[k] = v; wasText[k] = fv; }
      }
      const fill = p.fill && !(st.fill && st.fill.type === 'solid' && hexOf(st.fill.hex) === p.fill.hex && alphaOne(st.fill.alpha)) ? { hex: p.fill.hex } : null;
      const line = p.line && !(st.line && st.line.type === 'solid' && hexOf(st.line.hex) === p.line.hex && alphaOne(st.line.alpha)) ? { hex: p.line.hex } : null;
      const rec = { id, part: e.part, slideIndex: scene.index };
      if (Object.keys(text).length) rec.text = TEXT_KEYS.filter((k) => k in text).reduce((o, k) => ({ ...o, [k]: text[k] }), {});
      if (fill) rec.fill = fill; if (line) rec.line = line;
      if (!rec.text && !rec.fill && !rec.line) { const had = styles.delete(id); if (had) version++; return { ok: true, edit: null, changed: had }; }
      rec.was = {};
      if (rec.text) rec.was.text = TEXT_KEYS.filter((k) => k in rec.text).reduce((o, k) => ({ ...o, [k]: wasText[k] }), {});
      if (fill) rec.was.fill = { type: st.fill ? st.fill.type : 'none', hex: st.fill && st.fill.hex ? String(st.fill.hex).toUpperCase() : null, alpha: st.fill && st.fill.alpha != null ? st.fill.alpha : 1 };
      if (line) rec.was.line = { hex: st.line && st.line.hex ? String(st.line.hex).toUpperCase() : null, alpha: st.line && st.line.alpha != null ? st.line.alpha : 1 };
      const same = old && JSON.stringify(old) === JSON.stringify(rec);
      if (!same) { styles.set(id, rec); version++; }
      return { ok: true, edit: copySt(styles.get(id)), changed: !same };
    },
    /** A copy of the style record of an element, or null. */
    getStyle(id) { const r = styles.get(id); return r ? copySt(r) : null; },
    hasStyle(id) { return styles.has(id); },
    /** Copies of the style records of one slide part. */
    stylesForSlide(part) { return [...styles.values()].filter((r) => r.part === part).map(copySt); },
    listStyles() { return [...styles.values()].map(copySt); },
    /** Forget the style edit of one element (it has the file's style again; a move, a resize and a text edit of it stay). */
    discardStyle(id) { if (styles.delete(id)) version++; },
    /** The style (fill, line, shadow) of an element with its committed style edit applied (new objects; the scene is not written): see effectiveStyle. */
    effectiveStyle(e) { return effectiveStyle(e, e ? styles.get(e.id) : null); },
    /** The text of an element with its committed text edit AND its committed text style applied (new objects). `effectiveText(e)` above is unchanged: text edit only. */
    effectiveStyledText(e) { return effectiveText(e, e ? texts.get(e.id) : null, e ? styles.get(e.id) : null); },
    /** Copies of the resize records of one slide part. */
    resizesForSlide(part) { return [...resizes.values()].filter((r) => r.part === part).map(copyRz); },
    /** Copies of every resize record, in the order the elements were first resized. */
    listResizes() { return [...resizes.values()].map(copyRz); },
    /** Forget the resize of one element (it is back at its file size; a move of it stays). */
    discardResize(id) { if (resizes.delete(id)) version++; },
    /**
     * The own frame an element has once its committed move and resize are applied: { id, part, slideIndex, parentId, x, y, w, h, origin, moved, resized } in EMU in the parent space (x, y = its <a:off>;
     * w, h = its <a:ext>, or null when the element is not resized: the file's stays), or null when it has neither. Rotation and flips are not part of it: they stay as the file has them.
     */
    frame(id) {
      const m = moves.get(id); const r = resizes.get(id); if (!m && !r) return null;
      const base = r || m; const o = r ? r.origin : { x: m.origin.x, y: m.origin.y }; // a move-only element keeps its own <a:ext> (w, h null): the file has it, the record does not
      const w = r ? r.extent.w : null; const h = r ? r.extent.h : null;
      const dx = (m ? m.localDelta.dx : 0) + (r ? r.localShift.dx : 0); const dy = (m ? m.localDelta.dy : 0) + (r ? r.localShift.dy : 0);
      return { id, part: base.part, slideIndex: base.slideIndex, parentId: base.parentId, x: o.x + dx - (r ? (w - o.w) / 2 : 0), y: o.y + dy - (r ? (h - o.h) / 2 : 0), w, h, origin: { ...o }, moved: !!m, resized: !!r };
    },
    /** A copy of the record of an element, or null. */
    get(id) { const r = moves.get(id); return r ? copy(r) : null; },
    has(id) { return moves.has(id) || resizes.has(id) || texts.has(id) || styles.has(id); },
    /** Copies of the records of one slide part (e.g. 'ppt/slides/slide1.xml'). */
    forSlide(part) { return [...moves.values()].filter((r) => r.part === part).map(copy); },
    /** Copies of every record, in the order the elements were first moved. */
    list() { return [...moves.values()].map(copy); },
    // NEW TEXT BOXES (Phase 3.2.7-A). A record here is a text box that does not exist in the file: its frame and its text ARE the record (textBoxElement() turns it into the shape the scene walks), so the
    // box's own text edits replace `paras` and never make a text record; its moves, resizes and styles are ordinary records over its id. discard(id) and reset() leave the box itself (and its text) where it is.
    /** Make a new element on the slide of `scene`: an empty text box ('textbox') or a shape ('rect', 'roundRect', 'ellipse', 'line', 'arrow', Phase 3.2.7-B): { ok: true, id, box } with the record's copy. The xml id is above every id of the slide and of the elements already made on it. */
    addElement(scene, kind = 'textbox') {
      if (!ADDED_KINDS.includes(kind)) return { ok: false, reason: 'bad-kind' };
      if (!scene || !scene.byId || !scene.part || !(scene.width > 0) || !(scene.height > 0)) return { ok: false, reason: 'no-slide' };
      const used = [...scene.byId.keys()].map((k) => Number(String(k).split('#').pop())).filter(Number.isFinite); // (every element the scene walked, those made earlier included)
      const mine = [...added.values()].filter((r) => r.part === scene.part);
      const xmlId = Math.max(1, ...used, ...mine.map((r) => Number(r.xmlId))) + 1;
      const f = kind === 'textbox' ? defaultFrame(scene.width, scene.height, mine.length) : shapeFrame(kind, scene.width, scene.height, mine.length);
      const rec = { id: `${scene.part}#${xmlId}`, part: scene.part, xmlId, kind, name: `${KIND_LABEL[kind]} ${xmlId - 1}`, ...f, ...(kind === 'textbox' ? { paras: [{ text: '', rtl: false }] } : {}) };
      added.set(rec.id, rec); version++;
      return { ok: true, id: rec.id, box: copyAdded(rec) };
    },
    addTextBox(scene) { return this.addElement(scene, 'textbox'); },
    /** A text box made in the preview (not a shape): the ones with text of their own. */
    /** What typing into an EMPTY new text box looks like: the default run of its first paragraph with the style record applied ({ hex, font, rtl, sz, b, i }), or null. */
    addedLook(scene, id) {
      const e = scene && scene.byId ? scene.byId.get(id) : null; if (!e || !this.isAddedText(id) || !e.text) return null;
      const dr = effectiveText(e, null, styles.get(id))?.paras?.[0]?.defRun; if (!dr) return null;
      return { hex: dr.hex || '', font: dr.font, rtl: !!dr.rtl, sz: dr.sz || 0, b: !!dr.b, i: !!dr.i };
    },
    isAddedText(id) { const r = added.get(id); return !!r && r.kind === 'textbox'; },
    isAdded(id) { return added.has(id); },
    getAdded(id) { const r = added.get(id); return r ? copyAdded(r) : null; },
    /** The boxes made on one slide (by part), oldest first. */
    addedFor(part) { return [...added.values()].filter((r) => r.part === part).map(copyAdded); },
    listAdded() { return [...added.values()].map(copyAdded); },
    /** Replace the text of a box: `texts` = one string per paragraph, exactly as typed (a paragraph's direction follows its first strong letter). Returns { ok, changed, empty } (`empty`: nothing but white space is left). */
    setAddedText(id, texts) {
      const r = added.get(id); if (!r) return { ok: false, reason: 'unknown-box' }; if (r.kind !== 'textbox') return { ok: false, reason: 'not-a-text-box' };
      if (!Array.isArray(texts) || !texts.every((t) => typeof t === 'string')) return { ok: false, reason: 'bad-text' };
      const list = texts.length ? texts : ['']; const paras = list.map((text) => ({ text, rtl: paraRtl(text) }));
      const changed = JSON.stringify(paras) !== JSON.stringify(r.paras);
      if (changed) { r.paras = paras; version++; }
      return { ok: true, changed, empty: isBlankText(list) };
    },
    /** Delete a box with everything kept for it (its move, resize and style too). */
    removeAdded(id) { if (!added.delete(id)) return false; moves.delete(id); resizes.delete(id); texts.delete(id); styles.delete(id); version++; return true; },
    /** Forget every edit of one element, its move, its resize, its text and its style (it is back at its file frame, text and style). */
    discard(id) { const a = moves.delete(id); const b = resizes.delete(id); const c = texts.delete(id); const d = styles.delete(id); if (a || b || c || d) version++; },
    /** Forget every move, every resize, every text edit and every style edit, and every new text box (the deck is closing). */
    clear() { if (moves.size || resizes.size || texts.size || styles.size || added.size) version++; moves.clear(); resizes.clear(); texts.clear(); styles.clear(); added.clear(); },
    /** The number of elements that hold a committed move, a committed resize, a committed text edit or a committed style edit. */
    get size() { return new Set([...moves.keys(), ...resizes.keys(), ...texts.keys(), ...styles.keys()]).size; },
    get version() { return version; },
  };
}

const BY_DECK = new WeakMap();
/** The edit-state of one opened deck (the object parsePptx returned): one per deck for as long as that deck is open, gone with it. */
export function editStateOf(parsed) {
  let s = BY_DECK.get(parsed);
  if (!s) { s = createEditState(); BY_DECK.set(parsed, s); }
  return s;
}
