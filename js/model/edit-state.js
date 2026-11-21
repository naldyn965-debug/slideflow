// Temporary edit-state (Phase 3.2.3-B): the place where COMPLETED drag movements are kept. It is a small in-memory model that sits NEXT TO the Scene Model and never inside it:
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
// Why the local move is exact: an element's world matrix is W = M * Lm, with M = parent child space -> slide and Lm = the element's own box transform (translate to its centre, rotate,
// flip). Moving the box by l in the parent space gives W' = T(L * l) * W where L is the linear part of M (M's translation drops out). So the slide move d is reached by l = L^-1 * d: the
// group's rotation, flips and (non-uniform) scale are undone exactly, and the element's OWN rotation and flips never enter (they act about its centre, which only translates).
// M is taken from the scene as W * Lm^-1, so it holds for every element the scene placed, whatever put it there (own xfrm, inherited placeholder, nested groups, alternate content).
// A parent whose scale is zero (a collapsed group) has no inverse: such a move is refused ('singular') and the element stays where the file has it.
import { localMatrix, mul } from './scene.js';

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

/** A new, empty edit-state. */
export function createEditState() {
  const moves = new Map(); // element id -> record
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
    /** A copy of the record of an element, or null. */
    get(id) { const r = moves.get(id); return r ? copy(r) : null; },
    has(id) { return moves.has(id); },
    /** Copies of the records of one slide part (e.g. 'ppt/slides/slide1.xml'). */
    forSlide(part) { return [...moves.values()].filter((r) => r.part === part).map(copy); },
    /** Copies of every record, in the order the elements were first moved. */
    list() { return [...moves.values()].map(copy); },
    /** Forget the move of one element (it is back at its file position). */
    discard(id) { if (moves.delete(id)) version++; },
    /** Forget every move. */
    clear() { if (moves.size) version++; moves.clear(); },
    get size() { return moves.size; },
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
