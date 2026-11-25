// Element selection (Phase 3.2.1), with the move of the selected element on top of it (3.2.3-A preview, 3.2.3-B committed to a temporary edit-state, 3.2.3-C hardening) and the resize of the selected element (3.2.4-A preview, 3.2.4-B committed to the same edit-state); see the sections below.
// It never rotates, edits the text or saves anything, never touches the OOXML or the Scene Model, and adds nothing to the drawn slide until an element is selected (one outline <div> while a
// selection exists, removed again by clear(); the only other traces are a moved element's `translate` and a resized element's `transform`, both put back by reset() / dispose()).
//
// The selected element is identified by its stable Scene Model id (the same string as scene.byId keys and the `data-id` on the drawn element).
//
// What can be selected: every element of the SLIDE that is actually drawn (shape, picture, line, table, and the labelled boxes of charts / SmartArt / OLE).
//   - Groups are not selectable themselves: their children are drawn as separate elements and the child under the pointer is selected.
//   - Master / layout elements are drawn but inert (as in PowerPoint's slide view): clicking them counts as clicking empty space. Full-slide template
//     decoration would otherwise make "click empty space to clear" impossible on most real templates. One line decides this: `selectable()`.
//   - Hidden elements and the non-drawn branch of mc:AlternateContent are not drawn, so they cannot be selected.
//
// Hit-testing asks the browser what is painted at the pointer (document.elementsFromPoint: it already honours rotation, flips, clip paths and the stacking
// order = z-order of the slide). Only painted parts count: a shape's fill and outline, a picture inside its clip geometry, its text frame, its table, the
// labelled box of an undrawn element. The empty corners of an ellipse's frame, the inside of an outline-only shape and a diagonal line's bounding box do not,
// so the element that is visually on top at that point is the one selected. Hairline outlines (a 1 pt line is under 1 px wide) additionally get a small reach
// around the pointer (SLOP), measured against the real path, so lines and connectors can be hit with a mouse and a finger.
//
// While an element is selected its outline carries 8 resize handles and one rotation handle above the top-centre one (Phase 3.2.2-A / B). They are still transparent to input
// (pointer-events none, no listeners of their own, so what is under them is picked exactly as before); the 8 resize handles are found by GEOMETRY instead (3.2.4-A, below); the rotation handle is inert, nothing can be rotated.
//
// A selection is a click or a tap: press and release of the same primary pointer within TAP px, no cancel (a touch that turns into a scroll or pinch is
// cancelled by the browser, a mouse drag that selects text moves too far). Nothing is prevented or captured, so text can still be selected and copied.
//
// MOVE PREVIEW (Phase 3.2.3-A). A press on the ALREADY SELECTED element (the element that is visually on top at the pointer, hairline reach included) that travels more than TAP px
// drags it. The move is a TEMPORARY PREVIEW and nothing else: the Scene Model (scene.byId / e.world), the OOXML and storage are never touched, and nothing is saved.
//   - How: one CSS `translate` (the individual transform property) on the drawn element and on its outline, written through the CSSOM as `calc(var(--u) * dx) calc(var(--u) * dy)`.
//     `translate` is applied OUTSIDE the element's own `transform: rotate()`, so it is a pure move in slide space: rotation, flips (the drawn path is mirrored, not the frame) and group transforms
//     (a group child's `world` is already flattened into slide space) are untouched, and the outline and its 9 handles move with the element for free. The offset is kept in slide POINTS and
//     scaled by --u, so it stays exact at every zoom level, also when the slide is resized during or after the move.
//   - Accuracy: the pointer is converted to slide points (client px / px per point) on press and on every move, so the element follows the pointer exactly at any zoom, also when the page scrolls under a held mouse.
//   - The element's centre is kept inside the slide (a clamp), so an element can always be grabbed again; on its own this is a preview only (without an `edits` option, reset(), dispose() or drawing the slide again undoes it; with one, see 3.2.3-B below). There is no undo control in the UI yet.
//   - Cancel is safe: pointercancel, losing pointer capture, a second finger, a right/middle button, the window losing focus, the primary button released unseen, Escape, or select / clear / dispose mid-drag
//     all put the element back at its position from before this drag. Only a pointerup of the dragging pointer keeps the move.
//   - Touch: the selected element (only) gets `touch-action: none` so a finger drag on it moves it instead of scrolling the page (unselected elements and the slide itself keep `touch-action: auto`).
//     A finger that starts beside a hairline (outside its painted pixels) still scrolls the page and cancels, as before.
//   - Hit-testing follows the moved element: elementsFromPoint honours `translate`, and the hairline test subtracts the offset (nearOutline).
//
// COMMITTED MOVES (Phase 3.2.3-B). With an `edits` option (js/model/edit-state.js) a COMPLETED drag (the pointerup of the dragging pointer, nothing else) is committed to that temporary edit-state, which
// converts the slide-space offset to the parent group's local coordinates. The edit-state is the host's, not this controller's: dispose() only takes the preview off the drawn slide, the committed
// moves stay in the edit-state, and a controller attached to the same slide again (a redraw, another slide and back, another tab and back) draws them where they were left. Cancelling a drag
// never reaches the edit-state, so an earlier committed move is exactly what it was. Still nothing is written to the Scene Model or the OOXML, nothing is saved or exported. reset() is the
// explicit undo: it also forgets the committed moves of this slide. Without `edits` the controller behaves as in 3.2.3-A (a preview that ends with the controller).
//
// STRAY POINTERUP (Phase 3.2.3-C). Only the pointerup of the pointer that pressed counts, as with pointermove. A pointerup of any other pointer (another finger's or pen's release, a release that belongs to a press
// made elsewhere) is ignored: it neither drops a press that is still pending (the drag would never start) nor leaves an open drag without its release (the element would stop following and never be kept or
// committed). The release that does count still decides as before: drag kept (and committed), or click.
//
// RESIZE PREVIEW (Phase 3.2.4-A). A press within reach (HREACH, CSS px, by pointer type) of a visible resize handle of the SELECTED element, travelling more than TAP px, drags that handle. A press
// that does not travel is a click as before (it selects what is under the pointer). The handles stay transparent to input; handleAt() finds the nearest handle centre from their boxes, and
// skips the ones the small-frame rule hides. Like the move it is a TEMPORARY PREVIEW: the Scene Model, the OOXML, the edit-state and storage are never touched, nothing is saved or exported.
//   - Math: the pointer's travel since the press (slide points, so any zoom) is projected on the element's own rotated axes; the dragged edge(s) move by that much, the opposite edge / corner stays
//     (a corner moves both axes, an edge handle one; no aspect lock yet); the size stops at MIN_PT and at MAX_TIMES slide sizes; an axis the file gives under FLAT_PT (a line's zero height) is left alone.
//   - Drawing: the node is not re-laid out, it is STRETCHED: its `transform` becomes `translate(shift) <the rotation the renderer gave it> scale(sx, sy)` (scaled about its centre in its own frame, then
//     turned, then shifted in slide space; the original text is kept in tf0 and put back). Flips live inside the drawing (the svg), group transforms are already in the element's world frame, so both
//     are untouched. Text and pictures are stretched with it; PowerPoint would re-flow the text, which is for the step that commits a resize. The outline gets the resized frame itself (styleOutline)
//     and is not scaled, so the handles keep their screen size. The move `translate` is a separate property and is kept as it is; a move after a resize moves the stretched element.
//   - Cancel is the same as for a move (pointercancel, lost capture, a second finger, another button, blur, Escape, select / clear / dispose): the frame goes back to the size from before this drag, an earlier kept
//     (and committed) resize stays. Without an `edits` option a kept resize stays until reset() or dispose() and is not drawn again by the next controller (3.2.4-A); with one, see 3.2.4-B below.
//   - Cursor: hovering a handle sets the matching resize cursor on the slide and on the selected node (turned with the element's rotation). Touch: a finger that starts on a handle must not scroll the
//     page, and half of each handle is outside the element, so a non-passive `touchstart` listener cancels the scroll for touches that start within reach of a handle.
//   - Hit-testing follows the stretched element: elementsFromPoint honours the transform, the hairline test maps the pointer back to the file frame (nearOutline, sizeOf).
//
// COMMITTED RESIZES (Phase 3.2.4-B). With an `edits` option a COMPLETED resize (the pointerup of the dragging pointer, nothing else) is committed to the same temporary edit-state as a move: finishDrag() hands
// the frame the preview draws ({ w, h, tx, ty } in slide points) to edits.commitResize(), which converts it to the element's own <a:off> / <a:ext> in its PARENT coordinate space (a group's rotation, flips and
// non-uniform scale undone, through any number of nested groups) and checks the size limits again (js/model/resize-limits.js). If the edit-state refuses it the drag counts as cancelled, so the slide never shows a
// size the edit-state does not hold. A controller attached to the slide again (a redraw, another slide and back, another tab and back) draws the committed resizes before anything is selected, exactly as it
// draws the committed moves. Cancelling a resize never reaches the edit-state, so an earlier committed resize is exactly what it was. The drawing code (placeSize, styleOutline) is the preview code of 3.2.4-A
// and does not know the edit-state: only finishDrag() and the restore on attach do. Still nothing is written to the Scene Model or the OOXML (the scene keeps its file frames and `xfrm`s, the XML its original
// text), nothing is saved or exported. reset() is the explicit undo and also forgets the committed resizes of this slide; dispose() only takes the preview off the drawn slide.
//
// TEXT EDIT MODE (Phase 3.2.5-A; kept in the edit-state by 3.2.5-B). Double-clicking (a mouse; a pen or a finger: two taps within 350 ms and 24 px) the selected element, or beginEdit(), makes its drawn text editable in place when
// js/render/text-edit.js says the element is eligible (a text shape of the slide, no resize, no field, no vertical text, one column); otherwise `onEdit` reports why. While editing, this controller stands down
// for that element: no handles, no drag, no resize, the pointer belongs to the text (a press on it keeps the caret). A press anywhere else, select(), clear(), reset(), dispose() or Escape ends the mode and the
// ORIGINAL text DOM is put back; the draft is kept (handed to edits.commitText, the text box drawn again from the edit-state) on every exit except Escape, a structure violation and reset(), and is drawn again
// with the slide on every redraw, like the moves and resizes. Never written to the scene or the XML; this file asks text-edit.js, it does not touch the text itself.
import { MIN_PT, FLAT_PT, MAX_TIMES, SAME_PT } from '../model/resize-limits.js';
import { textEditability, startTextEdit } from './text-edit.js';
const E = 12700;
const SVG_NS = 'http://www.w3.org/2000/svg';
const HANDLES = ['nw', 'n', 'ne', 'e', 'se', 's', 'sw', 'w']; // Phase 3.2.2-A: the resize handles (positioned by css/render.css); they take no pointer events: a press on one is found by geometry (handleAt, Phase 3.2.4-A)
const TAP = { mouse: 4, pen: 6, touch: 10 }; // max pointer travel (CSS px) between press and release
const SLOP = { mouse: 3, pen: 6, touch: 10 }; // reach around the pointer (CSS px) for hairline outlines
const r3 = (n) => Math.round(n * 1000) / 1000;
const r6 = (n) => Math.round(n * 1e6) / 1e6;
const ZERO = Object.freeze({ dx: 0, dy: 0 });
// Resize (Phase 3.2.4-A). Handle -> direction of the dragged edge(s) in the element's own frame; reach (CSS px) around a handle's centre that a press must be within. The size limits (smallest size MIN_PT;
// an axis the file gives less than FLAT_PT (a line's zero height) cannot be resized by scaling and stays as it is; largest MAX_TIMES slide sizes) live in js/model/resize-limits.js, shared with the edit-state.
const DIRS = { nw: [-1, -1], n: [0, -1], ne: [1, -1], e: [1, 0], se: [1, 1], s: [0, 1], sw: [-1, 1], w: [-1, 0] };
const ANGLE = { e: 0, se: 45, s: 90, sw: 135, w: 180, nw: 225, n: 270, ne: 315 };
const CURSORS = ['ew-resize', 'nwse-resize', 'ns-resize', 'nesw-resize'];
const HREACH = { mouse: 7, pen: 10, touch: 16 };
/** Resize cursor for a handle on an element rotated by `rot` degrees (the handle turns with the element). */
const cursorFor = (k, rot) => CURSORS[((Math.round((((ANGLE[k] + (rot || 0)) % 360) + 360) % 360 / 45) % 8) + 8) % 4];
/** Scale factors [sx, sy] of a resize record against the file frame `W` (a flat axis is never scaled). */
const scaleOf = (W, r) => { const w = W.w / E; const h = W.h / E; return [r && w >= FLAT_PT ? r.w / w : 1, r && h >= FLAT_PT ? r.h / h : 1]; };

/** Ids that may be selected on this slide (see the header). */
const selectable = (scene, handle, id) => {
  const e = scene.byId.get(id); const node = handle.nodes.get(id);
  return !!e && !!node && !node.hidden && e.origin === 'slide' && e.kind !== 'group' && e.kind !== 'alt';
};

/** Slide elements painted at a client point, topmost first, one entry per element: [{ id, z }]. */
function paintedAt(root, handle, scene, x, y) {
  const out = []; const seen = new Set();
  for (const n of root.ownerDocument.elementsFromPoint(x, y)) {
    const owner = n.closest?.('.sf-el');
    if (!owner || !root.contains(owner)) continue;
    const geo = n.closest('.sf-geo');
    // the element's own box only counts for undrawn-content boxes; the bare <svg> box of a shape never does; anything inside (path, image, text, table cell) does
    const painted = n === owner ? owner.classList.contains('sf-opaque') : geo ? n !== geo : true;
    const id = owner.dataset.id;
    if (!painted || seen.has(id) || !selectable(scene, handle, id)) continue;
    seen.add(id); out.push({ id, z: Number(owner.dataset.z) });
  }
  return out;
}

/** Point in slide points -> the element's own frame (origin top-left, unrotated), undoing the rotation the renderer applied to the frame. */
function toFrame(W, px, py) {
  const t = ((W.rot || 0) * Math.PI) / 180; const c = Math.cos(t); const s = Math.sin(t);
  const dx = px - W.cx / E; const dy = py - W.cy / E;
  return [dx * c + dy * s + W.w / E / 2, -dx * s + dy * c + W.h / E / 2];
}

/** Highest element whose hairline outline (a path with an outline and no fill) passes within `reachPx` of the client point, or null. */
function nearOutline(root, handle, scene, x, y, reachPx, offsetOf = () => ZERO, sizeOf = () => null) {
  const r = root.getBoundingClientRect(); const u = r.width / (scene.width / E);
  if (!(u > 0)) return null;
  const px = (x - r.left) / u; const py = (y - r.top) / u; const reach = reachPx / u;
  let best = null; let probe = null;
  try {
    for (const [id, node] of handle.nodes) {
      if (!selectable(scene, handle, id)) continue;
      const e = scene.byId.get(id); const W = e.world;
      if (!W || (best && e.z < best.z)) continue;
      const off = offsetOf(id); // the element may have been moved in the preview: its world frame is where the scene put it, the pointer is tested against where it is drawn
      const rs = sizeOf(id); // ... and may have been resized in the preview: the drawn frame is the file frame scaled about its centre and shifted
      let [fx, fy] = toFrame(W, px - off.dx - (rs ? rs.tx : 0), py - off.dy - (rs ? rs.ty : 0));
      const w = W.w / E; const hh = W.h / E;
      if (rs) { const [kx, ky] = scaleOf(W, rs); fx = (fx - w / 2) / kx + w / 2; fy = (fy - hh / 2) / ky + hh / 2; }
      for (const p of node.querySelectorAll('.sf-geo > path')) {
        if (p.getAttribute('fill') !== 'none' || p.getAttribute('stroke') === 'none') continue;
        const sw = parseFloat(p.getAttribute('stroke-width')) || 0.25;
        const pad = reach + sw / 2 + 1;
        if (fx < -pad || fy < -pad || fx > w + pad || fy > hh + pad) continue;
        // The renderer mirrors the drawing (not the frame) for a horizontal flip: about the centre of its svg box, which is at least 1 pt wide.
        const lx = W.flipX ? Math.max(w, 1) - fx : fx;
        // Exact "within reach of this outline" test on a scratch copy of the path with a wider, undashed, round-capped stroke. The scratch <svg> lives only for
        // this call and is removed again; the drawn elements are not modified.
        if (!probe) { probe = document.createElementNS(SVG_NS, 'svg'); probe.setAttribute('width', '0'); probe.setAttribute('height', '0'); probe.style.position = 'absolute'; probe.style.visibility = 'hidden'; root.append(probe); }
        const q = document.createElementNS(SVG_NS, 'path');
        q.setAttribute('d', p.getAttribute('d')); q.setAttribute('fill', 'none'); q.setAttribute('stroke', '#000'); q.setAttribute('stroke-linecap', 'round'); q.setAttribute('stroke-linejoin', 'round');
        q.setAttribute('stroke-width', String(sw + 2 * reach));
        probe.replaceChildren(q);
        if (q.isPointInStroke(new DOMPoint(lx, fy))) { best = { id, z: e.z }; break; }
      }
    }
  } finally { probe?.remove(); }
  return best;
}

/**
 * Attach selection to one rendered slide. `handle` is what renderSlide() returned, `scene` the scene it drew.
 * Returns { id, pick(x, y, pointerType), select(id), clear(), offset(id), reset(), dragging, dispose() }. `onChange(id|null)` fires whenever the selection changes.
 * `onMove({ id, dx, dy, phase })` (Phase 3.2.3-A) fires when a drag starts ('start'), is kept ('end') or is undone ('cancel'); dx / dy are the element's preview offset in slide points.
 * `onEdit({ id, phase, reason?, message? })` (Phase 3.2.5-A) fires when text edit mode starts ('start'), ends ('end', with the reason: escape, outside, select, clear, reset, dispose, blur, structure, api) or is refused ('refused', with a reason code and a message for the status line).
 * `edits` (Phase 3.2.3-B / 3.2.4-B, optional) is the deck's temporary edit-state: completed moves and completed resizes are committed to it and its moves and resizes for this slide are drawn on attach (see the header).
 */
export function attachSelection(handle, scene, { onChange, onMove, onResize, onEdit, edits } = {}) {
  const { root } = handle;
  let selected = null; let outline = null; let down = null;
  const offsets = new Map(); // id -> { dx, dy } preview offset in slide points (never written to the scene)
  const sizes = new Map(); // id -> { w, h, tx, ty } resize (Phase 3.2.4-A preview, committed by 3.2.4-B): the frame as drawn, in slide points: size, and the shift of its centre from the file frame's (never written to the scene)
  const tf0 = new Map(); // id -> the `transform` the renderer gave the node (its rotation), put back when the resize preview goes
  let curNow = ''; // resize cursor currently shown
  let editing = null; // { id, session } while the text of the selected element is being edited (Phase 3.2.5-A)
  let lastTap = null; // { id, t, x, y } the last tap of a pen or finger on a selected element: a second one close by is a double tap
  let drag = null; // { id, pid, from: { dx, dy }, hdl, fromSize } while an element is being dragged (hdl: the handle of a resize drag, null for a move)
  const offsetOf = (id) => offsets.get(id) || ZERO;
  const sizeOf = (id) => sizes.get(id) || null;

  /** The id of the element that is visually on top at a client point, or null for empty space. */
  function pick(x, y, pointerType = 'mouse') {
    const type = SLOP[pointerType] ? pointerType : 'mouse';
    const top = paintedAt(root, handle, scene, x, y)[0] || null;
    const near = nearOutline(root, handle, scene, x, y, SLOP[type], offsetOf, sizeOf);
    const hit = near && (!top || near.z > top.z) ? near : top; // an outline within reach wins only when it is above what is exactly under the pointer
    return hit ? hit.id : null;
  }

  function drawOutline(id) {
    // Same frame as the element (copied from the style the renderer gave it, so rotation, scale and resizing follow automatically), at least 6 px thick so a
    // zero-height line still shows, kept centred on the element.
    const n = handle.nodes.get(id).style;
    const o = document.createElement('div'); o.className = 'sf-sel'; o.dataset.for = id; o.setAttribute('aria-hidden', 'true');
    styleOutline(o, id);
    const tr = n.getPropertyValue('translate'); if (tr) o.style.setProperty('translate', tr); // a moved element is selected again: the outline starts where it is drawn
    // 8 handles (corners + edge centres) as children of the outline: they sit at percentages of the same frame, so rotation, scale and zoom follow it for free
    const hs = document.createElement('div'); hs.className = 'sf-sel-h'; hs.setAttribute('aria-hidden', 'true');
    const rot = document.createElement('span'); rot.className = 'sf-rot'; hs.append(rot); // Phase 3.2.2-B: visual rotation handle above the top-centre handle (a stem and a knob; drawn first so the resize handle covers the stem's end)
    for (const k of HANDLES) { const hd = document.createElement('span'); hd.className = 'sf-hdl'; hd.dataset.h = k; hs.append(hd); }
    o.append(hs);
    root.append(o);
    return o;
  }

  /** Frame and rotation of the outline: the node's own (so a rotated or moved element is followed), or, when the element is resized in the preview, the resized frame (the node itself is only scaled by a
   *  `transform`, the outline is not, so its handles keep their screen size). */
  function styleOutline(o, id) {
    const n = handle.nodes.get(id).style; const r = sizes.get(id); const W = scene.byId.get(id)?.world;
    let left = n.left; let top = n.top; let width = n.width; let height = n.height; const tf = tf0.has(id) ? tf0.get(id) : n.transform;
    if (r && W) { const L2 = (v) => `calc(var(--u) * ${r3(v)})`; width = L2(r.w); height = L2(r.h); left = L2(W.cx / E + r.tx - r.w / 2); top = L2(W.cy / E + r.ty - r.h / 2); }
    const Wd = `max(${width}, 6px)`; const Ht = `max(${height}, 6px)`;
    o.style.width = Wd; o.style.height = Ht;
    o.style.left = `calc(${left} - (${Wd} - ${width}) / 2)`; o.style.top = `calc(${top} - (${Ht} - ${height}) / 2)`;
    if (tf) o.style.transform = tf; else o.style.removeProperty('transform');
  }

  // ---- move preview (Phase 3.2.3-A) ----
  /** Selected element: the cursor says it can be moved, and a finger that starts on it drags it instead of scrolling the page. Undone again when it is no longer selected. */
  function grabbable(id, on) {
    const n = handle.nodes.get(id); if (!n) return;
    if (on) { n.style.touchAction = 'none'; n.style.cursor = 'move'; } else { n.style.removeProperty('touch-action'); n.style.removeProperty('cursor'); }
  }
  /** Draw the preview offset of `id` on its node and on its outline (which has the same frame). No offset: no `translate` at all, so an unmoved element's DOM is untouched. */
  function place(id) {
    const o = offsetOf(id);
    const v = o.dx || o.dy ? `calc(var(--u) * ${r3(o.dx)}) calc(var(--u) * ${r3(o.dy)})` : '';
    for (const n of [handle.nodes.get(id), outline && selected === id ? outline : null]) { if (!n) continue; if (v) n.style.setProperty('translate', v); else n.style.removeProperty('translate'); }
  }
  function setOffset(id, o) { if (o.dx || o.dy) offsets.set(id, { dx: o.dx, dy: o.dy }); else offsets.delete(id); place(id); }
  /** Keep the element's centre inside the slide so it can always be grabbed again. An element whose centre is ALREADY outside when the drag starts (the file has it there, or a resize stretched it there) may not
   *  go further out, but it is not snapped back in either: on each axis the allowed range always contains the offset the drag started from `from` (Phase 3.2.4-C; before, a 10 pt drag could throw it 140 pt). */
  function clampOffset(id, o, from = ZERO) {
    const W = scene.byId.get(id)?.world; if (!W) return o;
    const rs = sizes.get(id); const sw = scene.width / E; const sh = scene.height / E; const cx = W.cx / E + (rs ? rs.tx : 0); const cy = W.cy / E + (rs ? rs.ty : 0); // the centre as drawn (a resized element's centre has shifted)
    const lim = (v, c, size, f) => Math.min(Math.max(size - c, f), Math.max(Math.min(-c, f), v));
    return { dx: lim(o.dx, cx, sw, from.dx), dy: lim(o.dy, cy, sh, from.dy) };
  }

  /** Client px -> slide points on the slide as it is drawn now (any zoom, any scroll), or null while the slide has no size (tab not showing). */
  function slidePt(x, y) {
    const r = root.getBoundingClientRect(); const u = r.width / (scene.width / E);
    return u > 0 ? [(x - r.left) / u, (y - r.top) / u] : null;
  }
  // ---- resize preview (Phase 3.2.4-A) ----
  /** The handle (key of DIRS) whose centre is within reach of a client point, nearest first, or null. A handle hidden by the small-frame rule has no box and is skipped. */
  function handleAt(x, y, type) {
    if (!outline) return null;
    let best = null; let bd = Infinity;
    for (const hd of outline.querySelectorAll('.sf-hdl')) {
      const r = hd.getBoundingClientRect(); if (!(r.width > 0)) continue;
      const dist = Math.hypot(x - (r.left + r.width / 2), y - (r.top + r.height / 2));
      if (dist <= HREACH[type] && dist < bd) { bd = dist; best = hd.dataset.h; }
    }
    return best;
  }
  function setCursor(v) { // the handles are transparent to input, so the resize cursor is set on the slide (and on the selected node, which has its own `move` cursor)
    if (v === curNow) return; curNow = v;
    if (v) root.style.cursor = v; else root.style.removeProperty('cursor');
    const n = selected != null ? handle.nodes.get(selected) : null; if (n) n.style.cursor = v || 'move';
  }
  const rotOf = (id) => scene.byId.get(id)?.world?.rot || 0;
  const frameOf = (id) => { const W = scene.byId.get(id).world; return sizes.get(id) || { w: W.w / E, h: W.h / E, tx: 0, ty: 0 }; };
  /** Draw the resize preview of `id`: its node gets `translate(shift) <its own rotation> scale(sx, sy)` as `transform` (scaled about its centre in its own frame, then turned, then shifted), its outline the resized frame. */
  function placeSize(id) {
    const n = handle.nodes.get(id); if (!n) return;
    const r = sizes.get(id); const W = scene.byId.get(id)?.world;
    if (r && W) {
      if (!tf0.has(id)) tf0.set(id, n.style.transform);
      const [kx, ky] = scaleOf(W, r);
      n.style.transform = `translate(calc(var(--u) * ${r3(r.tx)}), calc(var(--u) * ${r3(r.ty)})) ${tf0.get(id)} scale(${r6(kx)}, ${r6(ky)})`.replace(/\s+/g, ' ');
    } else if (tf0.has(id)) { const t = tf0.get(id); if (t) n.style.transform = t; else n.style.removeProperty('transform'); tf0.delete(id); }
    if (outline && selected === id) styleOutline(outline, id);
  }
  function setSize(id, r) {
    const W = scene.byId.get(id)?.world; if (!W) return;
    const same = !r || (Math.abs(r.w - W.w / E) < SAME_PT && Math.abs(r.h - W.h / E) < SAME_PT && Math.abs(r.tx) < SAME_PT && Math.abs(r.ty) < SAME_PT); // the edit-state's own test (SAME_PT): what is drawn is what it holds
    if (same) sizes.delete(id); else sizes.set(id, { w: r.w, h: r.h, tx: r.tx, ty: r.ty });
    placeSize(id);
  }
  const sizeInfo = (id) => { const W = scene.byId.get(id)?.world; const f = frameOf(id); return { w: f.w, h: f.h, dw: f.w - W.w / E, dh: f.h - W.h / E }; };
  /** Resize drag: the dragged edge(s) follow the pointer's travel since the press, measured in the element's own (rotated) frame; the opposite edge / corner stays where it was. */
  function resizeTo(p, d) {
    const W = scene.byId.get(drag.id)?.world; if (!W) return;
    const f = drag.fromSize || { w: W.w / E, h: W.h / E, tx: 0, ty: 0 };
    const [sx, sy] = DIRS[drag.hdl]; const t = (rotOf(drag.id) * Math.PI) / 180; const c = Math.cos(t); const s = Math.sin(t);
    const dx = p[0] - d.p0[0]; const dy = p[1] - d.p0[1];
    const fx = dx * c + dy * s; const fy = -dx * s + dy * c;
    const w = sx && W.w / E >= FLAT_PT ? Math.min(MAX_TIMES * (scene.width / E), Math.max(Math.min(MIN_PT, f.w), f.w + sx * fx)) : f.w;
    const h = sy && W.h / E >= FLAT_PT ? Math.min(MAX_TIMES * (scene.height / E), Math.max(Math.min(MIN_PT, f.h), f.h + sy * fy)) : f.h;
    const ax = (sx * (w - f.w)) / 2; const ay = (sy * (h - f.h)) / 2;
    setSize(drag.id, { w, h, tx: f.tx + ax * c - ay * s, ty: f.ty + ax * s + ay * c });
  }
  const blockSelect = (e) => e.preventDefault(); // no text selection while an element is being dragged
  const onKey = (e) => { if (e.key === 'Escape') { e.preventDefault(); e.stopPropagation(); finishDrag(false); } };
  const onBlur = () => finishDrag(false);
  function beginDrag(d) {
    if (drag || !selected || !d.p0) return false;
    drag = { id: selected, pid: d.id, from: { ...offsetOf(selected) }, hdl: d.hdl || null, fromSize: d.hdl ? sizeOf(selected) : null };
    if (drag.hdl) setCursor(cursorFor(drag.hdl, rotOf(selected)));
    try { root.setPointerCapture(d.id); } catch { /* the pointer is already gone: its pointerup / pointercancel will not come, the next event ends the drag */ }
    root.style.userSelect = 'none'; globalThis.getSelection?.()?.removeAllRanges();
    document.addEventListener('keydown', onKey, true); document.addEventListener('selectstart', blockSelect, true); window.addEventListener('blur', onBlur);
    if (drag.hdl) onResize?.({ id: drag.id, ...sizeInfo(drag.id), phase: 'start' }); else onMove?.({ id: drag.id, ...offsetOf(drag.id), phase: 'start' });
    return true;
  }
  /** Put the dragged element where the pointer is (relative to where it was pressed, both measured on the slide as drawn now). */
  function follow(e, d) {
    const p = drag && slidePt(e.clientX, e.clientY); if (!p) return;
    if (drag.hdl) { resizeTo(p, d); return; }
    setOffset(drag.id, clampOffset(drag.id, { dx: drag.from.dx + p[0] - d.p0[0], dy: drag.from.dy + p[1] - d.p0[1] }, drag.from));
  }
  /** End the drag. keep = the move / resize stays (pointerup); otherwise the element goes back to where / how it was before this drag. A kept move or resize is committed to the edit-state (when there is
   *  one); if that refuses it (a parent group that cannot be inverted, a frame outside the size limits) the drag counts as cancelled, so the drawn slide never shows a move or size the edit-state does not hold. */
  function finishDrag(keep) {
    if (!drag) return;
    const dr = drag; drag = null;
    if (dr.hdl) {
      if (keep && edits && !edits.commitResize(scene, dr.id, sizes.get(dr.id) || null).ok) keep = false; // no entry in `sizes` = back at the file size: that commit removes the record
      if (!keep) setSize(dr.id, dr.fromSize);
    } else {
      if (keep && edits) { const o = offsetOf(dr.id); if (!edits.commit(scene, dr.id, o.dx, o.dy).ok) keep = false; }
      if (!keep) setOffset(dr.id, dr.from);
    }
    setCursor('');
    try { root.releasePointerCapture(dr.pid); } catch { /* not captured any more */ }
    root.style.removeProperty('user-select');
    document.removeEventListener('keydown', onKey, true); document.removeEventListener('selectstart', blockSelect, true); window.removeEventListener('blur', onBlur);
    down = null;
    if (dr.hdl) onResize?.({ id: dr.id, ...sizeInfo(dr.id), phase: keep ? 'end' : 'cancel' }); else onMove?.({ id: dr.id, ...offsetOf(dr.id), phase: keep ? 'end' : 'cancel' });
  }

  /** Text edit mode (Phase 3.2.5-A): start it for the selected element (`pt` = client point for the caret), or report why not. Returns { ok: true } or { ok: false, reason, message }. */
  function beginEdit(id = selected, pt = null) {
    if (id == null || id !== selected) return { ok: false, reason: 'not-selected', message: 'Select a text shape first.' };
    if (editing) return { ok: true };
    finishDrag(false); setCursor('');
    const node = handle.nodes.get(id); const v = textEditability(scene.byId.get(id), { node, resized: !!sizes.get(id) });
    if (!v.ok) { onEdit?.({ id, phase: 'refused', reason: v.reason, message: v.message }); return v; }
    const e = scene.byId.get(id);
    const session = startTextEdit(node, { point: pt, onEnd: (why, draft) => afterEdit(id, why, draft), paras: (edits ? edits.effectiveText(e) : e.text).paras, orig: e.text.paras });
    if (!session) return { ok: false, reason: 'no-text', message: 'This shape has no drawn text to edit.' };
    editing = { id, session };
    node.style.touchAction = 'auto'; node.style.cursor = 'text'; // (the values are put back, in place, when the mode ends)
    outline?.classList.add('sf-editing');
    onEdit?.({ id, phase: 'start', locked: session.locked });
    return { ok: true };
  }
  /**
   * The session ended (by itself or through endEdit) and the DOM is the one it was drawn with: the controller is back to normal. `draft` = the text change to keep (Phase 3.2.5-B: every exit but Escape, a
   * structure violation and reset), handed to the edit-state, and the element's text box is drawn again from the edit-state so the slide shows what it holds. Without an edit-state nothing is kept (3.2.5-A).
   */
  function afterEdit(id, why, draft) {
    editing = null; let committed = false; let changed = false;
    if (draft && edits) {
      const r = edits.commitText(scene, id, draft); committed = r.ok; changed = r.ok && r.changed;
      if (changed) handle.setText?.(id, edits.effectiveText(scene.byId.get(id)), true, edits.hasText(id)); // (a shape back at the file's text is drawn as the file has it)
    }
    const node = handle.nodes.get(id);
    if (node && selected === id) { node.style.touchAction = 'none'; node.style.cursor = 'move'; }
    outline?.classList.remove('sf-editing');
    onEdit?.({ id, phase: 'end', reason: why, committed, changed, skipped: !!(draft && draft.skipped) });
  }
  /** End text edit mode (the original text is put back). */
  function endEdit(why = 'api') { if (editing) editing.session.end(why); }

  function select(id) {
    if (!selectable(scene, handle, id)) return false;
    if (id === selected) return true;
    endEdit('select');
    finishDrag(false); setCursor('');
    if (selected != null) grabbable(selected, false);
    outline?.remove();
    outline = drawOutline(id); selected = id; grabbable(id, true);
    onChange?.(id);
    return true;
  }
  function clear() {
    if (selected == null) return;
    endEdit('clear');
    finishDrag(false); setCursor('');
    grabbable(selected, false);
    outline?.remove(); outline = null; selected = null;
    onChange?.(null);
  }

  const typeOf = (e) => (e.pointerType === 'touch' || e.pointerType === 'pen' ? e.pointerType : 'mouse');
  const onDown = (e) => {
    finishDrag(false); // a new press while a drag is open (a second finger, another button) means the drag was not what the person meant
    if (editing) { // the text is being edited: a press on that element belongs to the text (caret, text selection), a press anywhere else ends the mode and is then an ordinary press
      if (editing.session.inner.contains(e.target) || handle.nodes.get(editing.id)?.contains(e.target) || pick(e.clientX, e.clientY, typeOf(e)) === selected) { down = null; return; } // (the blank part of an unfilled text box is not 'painted', but a press on it is a press on the shape)
      endEdit('outside');
    }
    if (!e.isPrimary || (e.pointerType === 'mouse' && e.button !== 0)) { down = null; return; } // a second finger or a non-primary button is not a click
    down = { id: e.pointerId, x: e.clientX, y: e.clientY, type: typeOf(e) };
    // pressed on the element that is already selected (and visually on top there): this press may become a drag
    const hd = selected ? handleAt(e.clientX, e.clientY, down.type) : null; // a handle of the selected element (they are transparent to input, so this is found by geometry): this press may become a resize
    if (hd) { down.p0 = slidePt(e.clientX, e.clientY); down.grab = !!down.p0; if (down.grab) down.hdl = hd; }
    else if (selected && pick(e.clientX, e.clientY, down.type) === selected) { down.p0 = slidePt(e.clientX, e.clientY); down.grab = !!down.p0; }
  };
  // a finger that starts on a handle resizes instead of scrolling the page (the handles are transparent to input and half of them lie outside the element, so touch-action cannot do this)
  const onTouchStart = (e) => { if (!editing && selected != null && e.touches.length === 1 && e.cancelable && handleAt(e.touches[0].clientX, e.touches[0].clientY, 'touch')) e.preventDefault(); };
  const onPointerMove = (e) => {
    const d = down;
    if (!d) { if (e.pointerType !== 'touch' && selected != null && !drag && !editing) { const k = handleAt(e.clientX, e.clientY, typeOf(e)); setCursor(k ? cursorFor(k, rotOf(selected)) : ''); } return; } // no press: the resize cursor over a handle
    if (!d.grab || e.pointerId !== d.id) return;
    if (d.type === 'mouse' && !(e.buttons & 1)) { // the button was released where we could not see it (a dialog, another window): finish like a release instead of sticking to the pointer
      if (drag) finishDrag(true); down = null; return;
    }
    if (!drag && Math.hypot(e.clientX - d.x, e.clientY - d.y) > TAP[d.type]) beginDrag(d);
    follow(e, d);
  };
  const onUp = (e) => {
    const d = down; if (!d || e.pointerId !== d.id) return; // a stray pointerup (no press of ours in flight, or another pointer's release) is not ours: it neither drops the pending press nor orphans an open drag
    down = null;
    if (d?.grab && e.pointerId === d.id) {
      if (!drag && Math.hypot(e.clientX - d.x, e.clientY - d.y) > TAP[d.type]) beginDrag(d); // released far from the press without a move event in between: still a drag
      if (drag) { follow(e, d); finishDrag(true); return; } // a drag is never also a click
    }
    if (!d || e.pointerId !== d.id || Math.hypot(e.clientX - d.x, e.clientY - d.y) > TAP[d.type]) return; // moved: a drag or a text selection, not a click
    const id = pick(d.x, d.y, d.type); // what was under the pointer when it was pressed
    if (id) {
      select(id);
      if (d.type !== 'mouse') { // a mouse double-clicks (onDblClick); a pen or a finger taps twice
        const now = e.timeStamp; const t = lastTap;
        if (t && t.id === id && now - t.t < 350 && Math.hypot(e.clientX - t.x, e.clientY - t.y) < 24) { lastTap = null; beginEdit(id, { x: e.clientX, y: e.clientY }); } else lastTap = { id, t: now, x: e.clientX, y: e.clientY };
      }
    } else { clear(); lastTap = null; }
  };
  const onDblClick = (e) => {
    if (editing || e.button !== 0) return;
    const id = pick(e.clientX, e.clientY, 'mouse'); if (id == null) return;
    if (id !== selected && !select(id)) return;
    beginEdit(id, { x: e.clientX, y: e.clientY });
  };
  const onCancel = () => { finishDrag(false); down = null; };
  const onLost = (e) => { if (drag && e.target === root && e.pointerId === drag.pid) finishDrag(false); }; // capture taken away before the pointer was released (a descendant losing implicit touch capture is not that)
  const onDragStart = (e) => { if (down?.grab || drag) e.preventDefault(); }; // the browser's own drag-and-drop of a text selection would cancel the pointer
  // moves committed earlier on this slide (a redraw, another slide and back, another tab and back) are drawn where they were left, before anything is selected
  if (edits) {
    for (const m of edits.forSlide(scene.part)) if (selectable(scene, handle, m.id)) setOffset(m.id, m.slideDelta);
    for (const m of edits.resizesForSlide(scene.part)) if (selectable(scene, handle, m.id)) setSize(m.id, m.slideFrame); // resizes committed earlier (3.2.4-B) come back the same way
    let texted = false; for (const m of edits.textsForSlide(scene.part)) if (selectable(scene, handle, m.id) && handle.setText?.(m.id, edits.effectiveText(scene.byId.get(m.id)), false, true)) texted = true; // and so do the texts (3.2.5-B)
    if (texted) handle.fitText();
  }
  root.addEventListener('pointerdown', onDown);
  root.addEventListener('touchstart', onTouchStart, { passive: false });
  root.addEventListener('pointermove', onPointerMove);
  root.addEventListener('pointerup', onUp);
  root.addEventListener('pointercancel', onCancel);
  root.addEventListener('lostpointercapture', onLost);
  root.addEventListener('dragstart', onDragStart);
  root.addEventListener('dblclick', onDblClick);

  return {
    get id() { return selected; },
    get dragging() { return drag ? drag.id : null; },
    /** The handle ('nw' … 'w') of the resize drag in progress, or null (no drag, or a move). */
    get resizing() { return drag ? drag.hdl : null; },
    pick, select, clear,
    /** The id of the element whose text is being edited (Phase 3.2.5-A), or null. */
    get editing() { return editing ? editing.id : null; },
    /** Start text edit mode for the selected element (see the header): { ok: true } or { ok: false, reason, message }. */
    beginEdit,
    /** End text edit mode: the original text is put back. */
    endEdit,
    /** The draft text of the element being edited, one string per paragraph (null when not editing). Never written anywhere. */
    editText() { return editing ? editing.session.text() : null; },
    /** The preview offset of an element in slide points ({ dx: 0, dy: 0 } when it has not been moved). */
    offset(id) { return { ...offsetOf(id) }; },
    /** Put every moved and resized element of this slide back at its file frame, committed moves and resizes included (the explicit undo: there is nothing to undo in the scene). */
    reset() { endEdit('reset'); finishDrag(false); if (edits) { let t = false; for (const m of edits.textsForSlide(scene.part)) { edits.discardText(m.id); handle.setText?.(m.id, scene.byId.get(m.id).text, false); t = true; } if (t) handle.fitText(); } for (const id of [...sizes.keys()]) { setSize(id, null); edits?.discardResize(id); } for (const id of [...offsets.keys()]) { setOffset(id, ZERO); edits?.discard(id); } },
    /** The size of an element as drawn, in slide points: { w, h } and { dw, dh } against the file ({ dw: 0, dh: 0 } when it has not been resized). With an `edits` option a completed resize is in the edit-state too. */
    size(id) { return sizeInfo(id); },
    dispose() {
      endEdit('dispose'); finishDrag(false); setCursor(''); for (const id of [...sizes.keys()]) setSize(id, null); for (const id of [...offsets.keys()]) setOffset(id, ZERO); // the preview goes off the drawn slide; committed moves and resizes stay in the edit-state and are drawn again by the next controller
      root.removeEventListener('pointerdown', onDown); root.removeEventListener('touchstart', onTouchStart); root.removeEventListener('pointermove', onPointerMove); root.removeEventListener('pointerup', onUp); root.removeEventListener('pointercancel', onCancel);
      root.removeEventListener('lostpointercapture', onLost); root.removeEventListener('dragstart', onDragStart); root.removeEventListener('dblclick', onDblClick);
      if (selected != null) grabbable(selected, false);
      outline?.remove(); outline = null; selected = null; down = null;
    },
  };
}
