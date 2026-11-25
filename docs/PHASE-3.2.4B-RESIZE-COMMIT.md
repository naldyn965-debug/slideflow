# Phase 3.2.4-B: commit a completed resize to the in-memory edit-state

A **completed** resize (the `pointerup` of the dragging pointer) is now kept in the same temporary, in-memory edit-state as the committed moves of 3.2.3-B. It survives redraws, other slides and other tabs for as long as the deck is open. Still **no Scene Model change, no PPTX / XML edit, nothing saved or exported**: the edit-state sits next to the Scene Model, never inside it, and dies with the opened deck.

## Behaviour
- **Kept**: a resize released by the dragging pointer. Resizing back to exactly the file size removes the record. A second resize builds on the kept size and the new record is the **total** since the file.
- **Drawn again**: a controller attached to a slide (redraw, another slide and back, another tab and back) draws the committed resizes (and moves) before anything is selected. Selecting such an element shows its outline and handles on the resized frame; the status line reports the kept size as before.
- **Cancelled**: `pointercancel`, lost capture, a second finger, another button, blur, Escape, select / clear / dispose in mid-drag never reach the edit-state. The frame returns to the *earlier committed* size (or the file's), the record and the version are byte-identical to before.
- **Refused**: if the edit-state refuses the frame the release counts as a cancel (`start`, `cancel`), so the slide never shows a size the model does not hold.
- **Undo**: `reset()` forgets the committed moves and resizes of the slide; `dispose()` only takes the preview off the drawn slide.

## What is stored (`js/model/edit-state.js`, separate map, move records unchanged)
`commitResize(scene, id, { w, h, tx, ty })` takes the frame the preview draws (slide points: size in the element's own rotated axes, shift of its centre) and keeps:

| field | meaning |
|---|---|
| `slideFrame` | `{ w, h, tx, ty }` as drawn, slide points (exact at every zoom) |
| `scale` | `{ sx, sy }` against the file's world frame (a flat axis is 1) |
| `origin` | the file's own `<a:off>` / `<a:ext>` `{ x, y, w, h }` (EMU, parent space) |
| `extent` | the resized `<a:ext>` `{ w, h }` = `origin.w * sx`, `origin.h * sy` |
| `localShift` | centre shift in the **parent** space (EMU) |
| `position` | the resized `<a:off>` when not moved: `origin + localShift - (extent - origin) / 2` |
| `id, part, slideIndex, parentId, kind, posSource` | as for a move |

`frame(id)` composes a move and a resize of one element into the `<a:off>` / `<a:ext>` its own xfrm would hold (both are centre shifts and add up). Other additions: `getResize`, `hasResize`, `resizesForSlide`, `listResizes`, `discardResize`. `has()`, `size` (elements with a move or a resize), `version`, `discard()` (move **and** resize) and `clear()` cover both; `list()`, `forSlide()`, `get()` are still the moves only.

## Why the conversion is exact
The preview stretches the frame about its centre in its own axes. The element's own rotation and flips act about its centre and do not move it or its axes; the parent's child space only scales the box. So scaling the own extent by `(sx, sy)` scales the drawn frame by the same factors, and the centre goes where the slide shift says through the same `L^-1` as a move (the group's rotation, flips and non-uniform scale undone, through any number of nested groups). Rotation and flips stay in the file as they are. A parent that cannot be inverted (collapsed group) is refused (`singular`), as for a move.

## Limits (`js/model/resize-limits.js`, one place, used by the drag and the edit-state)
2 pt minimum on a scaled axis (an axis the file already has under 2 pt but not flat keeps its size or grows), four slide widths / heights maximum, an axis under 1 pt (a line's zero height) is never scaled. The edit-state checks them again, so a frame the drag could not have produced is refused (`too-small`, `too-large`, `flat-axis`, `invalid`).

## Limitations (on purpose)
- Values are floats in EMU, as for moves; rounding to whole EMU belongs to the step that writes the file.
- Text and pictures are still **stretched** in the drawing; PowerPoint re-flows text. Re-flowing / re-fitting belongs to the step that writes the file.
- A child of a group with a non-uniform scale under a rotation (shear, already reported "approximated") is committed against the same unsheared frame the renderer draws.
- An element whose xfrm is inherited from a layout (`posSource` not `own`) is recorded against that inherited frame; writing a new xfrm for it is for the export step.
- No aspect lock, no keyboard resize, no undo control in the UI, rotation handle still inert.

## Tests
`node tests/selection.mjs` (sections "resize commit (Phase 3.2.4-B)"), `render.mjs`, `e2e.mjs`. See TESTING.md.
