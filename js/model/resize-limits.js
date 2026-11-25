// Size limits of a resize (Phase 3.2.4-A preview, 3.2.4-B commit), in slide POINTS. One place, no imports, so the drag preview (js/render/selection.js) and the edit-state that keeps
// a completed resize (js/model/edit-state.js) can never disagree about what a valid frame is.
//   MIN_PT     the smallest a dragged frame gets (an axis the file already has smaller, but not flat, keeps its file size instead: a drag never makes it bigger just to reach the minimum)
//   FLAT_PT    an axis the file gives under this (a line's zero height) is not scaled at all: scaling nothing gives nothing, and the axis must stay exactly as it is
//   MAX_TIMES  the largest frame, in slide widths / heights
export const MIN_PT = 2;
export const FLAT_PT = 1;
export const MAX_TIMES = 4;
// A frame within SAME_PT of the file's (size and centre shift) IS the file's frame: the drag preview (nothing drawn) and the edit-state (no record) must agree on this one number, or a release a hair away
// from the file frame would leave a stretched element on the slide that the edit-state does not hold (Phase 3.2.4-C).
export const SAME_PT = 1e-6;
