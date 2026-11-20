// Element selection (Phase 3.2.1). SELECTION ONLY: it never moves, resizes, edits or saves anything, never touches the OOXML, and adds nothing to the drawn
// slide until an element is selected (one outline <div> while a selection exists, removed again by clear()).
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
// While an element is selected its outline carries 8 resize handles and one rotation handle above the top-centre one (Phase 3.2.2-A / B). They are VISUAL ONLY: pointer-events none, no cursor, no listeners, so selection and
// touch behave exactly as before and nothing can be resized or rotated.
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
//   - The element's centre is kept inside the slide (a clamp), so an element can always be grabbed again; this is a preview only, there is no undo control yet (reset() or drawing the slide again undoes it).
//   - Cancel is safe: pointercancel, losing pointer capture, a second finger, a right/middle button, the window losing focus, the primary button released unseen, Escape, or select / clear / dispose mid-drag
//     all put the element back at its position from before this drag. Only a pointerup of the dragging pointer keeps the move.
//   - Touch: the selected element (only) gets `touch-action: none` so a finger drag on it moves it instead of scrolling the page (unselected elements and the slide itself keep `touch-action: auto`).
//     A finger that starts beside a hairline (outside its painted pixels) still scrolls the page and cancels, as before.
//   - Hit-testing follows the moved element: elementsFromPoint honours `translate`, and the hairline test subtracts the offset (nearOutline).
const E = 12700;
const SVG_NS = 'http://www.w3.org/2000/svg';
const HANDLES = ['nw', 'n', 'ne', 'e', 'se', 's', 'sw', 'w']; // Phase 3.2.2-A: purely visual resize handles (positioned by css/render.css); they take no pointer events and do nothing yet
const TAP = { mouse: 4, pen: 6, touch: 10 }; // max pointer travel (CSS px) between press and release
const SLOP = { mouse: 3, pen: 6, touch: 10 }; // reach around the pointer (CSS px) for hairline outlines
const r3 = (n) => Math.round(n * 1000) / 1000;
const ZERO = Object.freeze({ dx: 0, dy: 0 });

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
function nearOutline(root, handle, scene, x, y, reachPx, offsetOf = () => ZERO) {
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
      const [fx, fy] = toFrame(W, px - off.dx, py - off.dy);
      const w = W.w / E; const hh = W.h / E;
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
 * `onMove({ id, dx, dy, phase })` (Phase 3.2.3-A, preview only) fires when a drag starts ('start'), is kept ('end') or is undone ('cancel'); dx / dy are the element's preview offset in slide points.
 */
export function attachSelection(handle, scene, { onChange, onMove } = {}) {
  const { root } = handle;
  let selected = null; let outline = null; let down = null;
  const offsets = new Map(); // id -> { dx, dy } preview offset in slide points (never written to the scene)
  let drag = null; // { id, pid, from: { dx, dy } } while an element is being dragged
  const offsetOf = (id) => offsets.get(id) || ZERO;

  /** The id of the element that is visually on top at a client point, or null for empty space. */
  function pick(x, y, pointerType = 'mouse') {
    const type = SLOP[pointerType] ? pointerType : 'mouse';
    const top = paintedAt(root, handle, scene, x, y)[0] || null;
    const near = nearOutline(root, handle, scene, x, y, SLOP[type], offsetOf);
    const hit = near && (!top || near.z > top.z) ? near : top; // an outline within reach wins only when it is above what is exactly under the pointer
    return hit ? hit.id : null;
  }

  function drawOutline(id) {
    // Same frame as the element (copied from the style the renderer gave it, so rotation, scale and resizing follow automatically), at least 6 px thick so a
    // zero-height line still shows, kept centred on the element.
    const n = handle.nodes.get(id).style;
    const o = document.createElement('div'); o.className = 'sf-sel'; o.dataset.for = id; o.setAttribute('aria-hidden', 'true');
    const W = `max(${n.width}, 6px)`; const H = `max(${n.height}, 6px)`;
    o.style.width = W; o.style.height = H;
    o.style.left = `calc(${n.left} - (${W} - ${n.width}) / 2)`; o.style.top = `calc(${n.top} - (${H} - ${n.height}) / 2)`;
    if (n.transform) o.style.transform = n.transform;
    const tr = n.getPropertyValue('translate'); if (tr) o.style.setProperty('translate', tr); // a moved element is selected again: the outline starts where it is drawn
    // 8 handles (corners + edge centres) as children of the outline: they sit at percentages of the same frame, so rotation, scale and zoom follow it for free
    const hs = document.createElement('div'); hs.className = 'sf-sel-h'; hs.setAttribute('aria-hidden', 'true');
    const rot = document.createElement('span'); rot.className = 'sf-rot'; hs.append(rot); // Phase 3.2.2-B: visual rotation handle above the top-centre handle (a stem and a knob; drawn first so the resize handle covers the stem's end)
    for (const k of HANDLES) { const hd = document.createElement('span'); hd.className = 'sf-hdl'; hd.dataset.h = k; hs.append(hd); }
    o.append(hs);
    root.append(o);
    return o;
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
  /** Keep the element's centre inside the slide so it can always be grabbed again. */
  function clampOffset(id, o) {
    const W = scene.byId.get(id)?.world; if (!W) return o;
    const sw = scene.width / E; const sh = scene.height / E; const cx = W.cx / E; const cy = W.cy / E;
    return { dx: Math.min(sw - cx, Math.max(-cx, o.dx)), dy: Math.min(sh - cy, Math.max(-cy, o.dy)) };
  }
  /** Client px -> slide points on the slide as it is drawn now (any zoom, any scroll), or null while the slide has no size (tab not showing). */
  function slidePt(x, y) {
    const r = root.getBoundingClientRect(); const u = r.width / (scene.width / E);
    return u > 0 ? [(x - r.left) / u, (y - r.top) / u] : null;
  }
  const blockSelect = (e) => e.preventDefault(); // no text selection while an element is being dragged
  const onKey = (e) => { if (e.key === 'Escape') { e.preventDefault(); e.stopPropagation(); finishDrag(false); } };
  const onBlur = () => finishDrag(false);
  function beginDrag(d) {
    if (drag || !selected || !d.p0) return false;
    drag = { id: selected, pid: d.id, from: { ...offsetOf(selected) } };
    try { root.setPointerCapture(d.id); } catch { /* the pointer is already gone: its pointerup / pointercancel will not come, the next event ends the drag */ }
    root.style.userSelect = 'none'; globalThis.getSelection?.()?.removeAllRanges();
    document.addEventListener('keydown', onKey, true); document.addEventListener('selectstart', blockSelect, true); window.addEventListener('blur', onBlur);
    onMove?.({ id: drag.id, ...offsetOf(drag.id), phase: 'start' });
    return true;
  }
  /** Put the dragged element where the pointer is (relative to where it was pressed, both measured on the slide as drawn now). */
  function follow(e, d) {
    const p = drag && slidePt(e.clientX, e.clientY); if (!p) return;
    setOffset(drag.id, clampOffset(drag.id, { dx: drag.from.dx + p[0] - d.p0[0], dy: drag.from.dy + p[1] - d.p0[1] }));
  }
  /** End the drag. keep = the move stays (pointerup); otherwise the element goes back to where it was before this drag. */
  function finishDrag(keep) {
    if (!drag) return;
    const dr = drag; drag = null;
    if (!keep) setOffset(dr.id, dr.from);
    try { root.releasePointerCapture(dr.pid); } catch { /* not captured any more */ }
    root.style.removeProperty('user-select');
    document.removeEventListener('keydown', onKey, true); document.removeEventListener('selectstart', blockSelect, true); window.removeEventListener('blur', onBlur);
    down = null;
    onMove?.({ id: dr.id, ...offsetOf(dr.id), phase: keep ? 'end' : 'cancel' });
  }

  function select(id) {
    if (!selectable(scene, handle, id)) return false;
    if (id === selected) return true;
    finishDrag(false);
    if (selected != null) grabbable(selected, false);
    outline?.remove();
    outline = drawOutline(id); selected = id; grabbable(id, true);
    onChange?.(id);
    return true;
  }
  function clear() {
    if (selected == null) return;
    finishDrag(false);
    grabbable(selected, false);
    outline?.remove(); outline = null; selected = null;
    onChange?.(null);
  }

  const typeOf = (e) => (e.pointerType === 'touch' || e.pointerType === 'pen' ? e.pointerType : 'mouse');
  const onDown = (e) => {
    finishDrag(false); // a new press while a drag is open (a second finger, another button) means the drag was not what the person meant
    if (!e.isPrimary || (e.pointerType === 'mouse' && e.button !== 0)) { down = null; return; } // a second finger or a non-primary button is not a click
    down = { id: e.pointerId, x: e.clientX, y: e.clientY, type: typeOf(e) };
    // pressed on the element that is already selected (and visually on top there): this press may become a drag
    if (selected && pick(e.clientX, e.clientY, down.type) === selected) { down.p0 = slidePt(e.clientX, e.clientY); down.grab = !!down.p0; }
  };
  const onPointerMove = (e) => {
    const d = down; if (!d || !d.grab || e.pointerId !== d.id) return;
    if (d.type === 'mouse' && !(e.buttons & 1)) { // the button was released where we could not see it (a dialog, another window): finish like a release instead of sticking to the pointer
      if (drag) finishDrag(true); down = null; return;
    }
    if (!drag && Math.hypot(e.clientX - d.x, e.clientY - d.y) > TAP[d.type]) beginDrag(d);
    follow(e, d);
  };
  const onUp = (e) => {
    const d = down; down = null;
    if (d?.grab && e.pointerId === d.id) {
      if (!drag && Math.hypot(e.clientX - d.x, e.clientY - d.y) > TAP[d.type]) beginDrag(d); // released far from the press without a move event in between: still a drag
      if (drag) { follow(e, d); finishDrag(true); return; } // a drag is never also a click
    }
    if (!d || e.pointerId !== d.id || Math.hypot(e.clientX - d.x, e.clientY - d.y) > TAP[d.type]) return; // moved: a drag or a text selection, not a click
    const id = pick(d.x, d.y, d.type); // what was under the pointer when it was pressed
    if (id) select(id); else clear();
  };
  const onCancel = () => { finishDrag(false); down = null; };
  const onLost = (e) => { if (drag && e.target === root && e.pointerId === drag.pid) finishDrag(false); }; // capture taken away before the pointer was released (a descendant losing implicit touch capture is not that)
  const onDragStart = (e) => { if (down?.grab || drag) e.preventDefault(); }; // the browser's own drag-and-drop of a text selection would cancel the pointer
  root.addEventListener('pointerdown', onDown);
  root.addEventListener('pointermove', onPointerMove);
  root.addEventListener('pointerup', onUp);
  root.addEventListener('pointercancel', onCancel);
  root.addEventListener('lostpointercapture', onLost);
  root.addEventListener('dragstart', onDragStart);

  return {
    get id() { return selected; },
    get dragging() { return drag ? drag.id : null; },
    pick, select, clear,
    /** The preview offset of an element in slide points ({ dx: 0, dy: 0 } when it has not been moved). */
    offset(id) { return { ...offsetOf(id) }; },
    /** Put every moved element back (the preview only: there was nothing to undo in the scene). */
    reset() { finishDrag(false); for (const id of [...offsets.keys()]) setOffset(id, ZERO); },
    dispose() {
      finishDrag(false); for (const id of [...offsets.keys()]) setOffset(id, ZERO); // the preview moves go with the controller: nothing is left on the drawn slide
      root.removeEventListener('pointerdown', onDown); root.removeEventListener('pointermove', onPointerMove); root.removeEventListener('pointerup', onUp); root.removeEventListener('pointercancel', onCancel);
      root.removeEventListener('lostpointercapture', onLost); root.removeEventListener('dragstart', onDragStart);
      if (selected != null) grabbable(selected, false);
      outline?.remove(); outline = null; selected = null; down = null;
    },
  };
}
