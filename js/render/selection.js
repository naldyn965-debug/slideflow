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
// A selection is a click or a tap: press and release of the same primary pointer within TAP px, no cancel (a touch that turns into a scroll or pinch is
// cancelled by the browser, a mouse drag that selects text moves too far). Nothing is prevented or captured, so text can still be selected and copied.
const E = 12700;
const SVG_NS = 'http://www.w3.org/2000/svg';
const TAP = { mouse: 4, pen: 6, touch: 10 }; // max pointer travel (CSS px) between press and release
const SLOP = { mouse: 3, pen: 6, touch: 10 }; // reach around the pointer (CSS px) for hairline outlines

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
function nearOutline(root, handle, scene, x, y, reachPx) {
  const r = root.getBoundingClientRect(); const u = r.width / (scene.width / E);
  if (!(u > 0)) return null;
  const px = (x - r.left) / u; const py = (y - r.top) / u; const reach = reachPx / u;
  let best = null; let probe = null;
  try {
    for (const [id, node] of handle.nodes) {
      if (!selectable(scene, handle, id)) continue;
      const e = scene.byId.get(id); const W = e.world;
      if (!W || (best && e.z < best.z)) continue;
      const [fx, fy] = toFrame(W, px, py);
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
 * Returns { id, pick(x, y, pointerType), select(id), clear(), dispose() }. `onChange(id|null)` fires whenever the selection changes.
 */
export function attachSelection(handle, scene, { onChange } = {}) {
  const { root } = handle;
  let selected = null; let outline = null; let down = null;

  /** The id of the element that is visually on top at a client point, or null for empty space. */
  function pick(x, y, pointerType = 'mouse') {
    const type = SLOP[pointerType] ? pointerType : 'mouse';
    const top = paintedAt(root, handle, scene, x, y)[0] || null;
    const near = nearOutline(root, handle, scene, x, y, SLOP[type]);
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
    root.append(o);
    return o;
  }

  function select(id) {
    if (!selectable(scene, handle, id)) return false;
    if (id === selected) return true;
    outline?.remove();
    outline = drawOutline(id); selected = id;
    onChange?.(id);
    return true;
  }
  function clear() {
    if (selected == null) return;
    outline?.remove(); outline = null; selected = null;
    onChange?.(null);
  }

  const typeOf = (e) => (e.pointerType === 'touch' || e.pointerType === 'pen' ? e.pointerType : 'mouse');
  const onDown = (e) => {
    if (!e.isPrimary || (e.pointerType === 'mouse' && e.button !== 0)) { down = null; return; } // a second finger or a non-primary button is not a click
    down = { id: e.pointerId, x: e.clientX, y: e.clientY, type: typeOf(e) };
  };
  const onUp = (e) => {
    const d = down; down = null;
    if (!d || e.pointerId !== d.id || Math.hypot(e.clientX - d.x, e.clientY - d.y) > TAP[d.type]) return; // moved: a drag or a text selection, not a click
    const id = pick(d.x, d.y, d.type); // what was under the pointer when it was pressed
    if (id) select(id); else clear();
  };
  const onCancel = () => { down = null; };
  root.addEventListener('pointerdown', onDown);
  root.addEventListener('pointerup', onUp);
  root.addEventListener('pointercancel', onCancel);

  return {
    get id() { return selected; },
    pick, select, clear,
    dispose() {
      root.removeEventListener('pointerdown', onDown); root.removeEventListener('pointerup', onUp); root.removeEventListener('pointercancel', onCancel);
      outline?.remove(); outline = null; selected = null; down = null;
    },
  };
}
