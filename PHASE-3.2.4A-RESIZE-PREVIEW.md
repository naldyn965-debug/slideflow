# Phase 3.2.4-A: drag the 8 resize handles (temporary resize preview)

Press a resize handle of the **selected** element and drag: the element is stretched to follow the pointer. Like the move of 3.2.3-A it is a **temporary preview** and nothing else: no Scene Model change, no PPTX / XML edit, **no edit-state commit**, nothing saved or exported. No rotation, no aspect lock, no keyboard resize.

## Behaviour
- **What starts a resize**: a press within reach of a visible handle of the selected element (mouse 7 px, pen 10 px, touch 16 px around the handle's centre; the nearest handle wins), followed by travel of more than the tap limit (mouse 4 px, pen 6 px, touch 10 px). A press on a handle that does not travel is a click as before: it selects what is under the pointer (the handles are transparent to input and always were). A handle hidden by the small-frame rule (very small frames) cannot be pressed.
- **Which edges move**: a corner handle moves its two edges, an edge handle one. The **opposite** edge / corner stays exactly where it was, in the element's own (rotated) frame: the pointer's travel is projected on the element's rotated axes. The size stops at 2 pt (an edge dragged past the other does not flip) and at four slide sizes. An axis the file gives under 1 pt (a line's zero height) is left alone, so a flat line widens but never gets tall.
- **How long it lasts**: until `reset()` or `dispose()` on the controller. A resize is not committed anywhere and is not drawn again by the next controller (another slide and back, another tab, a redraw): the slide shows the file again. The committed moves of 3.2.3-B are not affected: they come back as before.
- **Cancel** is the same as for a move: `pointercancel`, losing pointer capture, a second finger, another mouse button, the window losing focus, the primary button released unseen, Escape, select / clear / dispose in mid-drag. The frame goes back to the size from before **this** drag; an earlier kept resize stays. A `pointerup` of another pointer is ignored (3.2.3-C). Only the release of the dragging pointer keeps the resize.
- **Cursor**: hovering a handle shows the matching resize cursor (`ew` / `ns` / `nwse` / `nesw`, turned with the element's rotation) on the slide and on the selected node (whose own cursor is `move`). Touch pointers do not change the cursor.
- **Touch**: a finger that starts within reach of a handle resizes instead of scrolling the page. Half of every handle lies outside the element, where `touch-action: none` cannot reach, so a non-passive `touchstart` listener cancels the scroll for such touches. A finger anywhere else scrolls as before.
- **Status line** (Rendered tab): "Selected: Shape “name” (id). Resized in the preview to 360 × 240 pt (not saved)", in slide points, after "Moved in the preview by ..." when the element is moved too.
- **Move and resize together**: a resized element can be moved (it moves as drawn, the edit-state record holds the move only) and a moved element can be resized (the frame grows from its moved position).

## How (all in `js/render/selection.js`, plus the status line in `js/ui/rendered.js`)
- **The node is stretched, not re-laid out.** Its `transform` becomes `translate(shift) <the rotation the renderer gave it> scale(sx, sy)`: scaled about its centre in its own frame, then turned, then shifted in slide space. The original `transform` text is kept and written back verbatim when the preview goes. The node's `left / top / width / height`, the drawing inside it and its flips (the svg is mirrored, not the frame) and the group transforms (a group child's frame is already flattened into slide space) are not touched.
- **The outline is the resized frame**, set through the CSSOM from the same numbers (`styleOutline`), and is not scaled, so the handles keep their screen size and stay on the frame. The move `translate` is a separate property and is left as it is.
- **Points, not pixels**: the preview is kept as `{ w, h, tx, ty }` in slide points (size, and the shift of the frame's centre) and written as `calc(var(--u) * n)`, so it is exact at every zoom and follows `setWidth()` afterwards. The pointer is converted to slide points on press and on every move, as for a move.
- **Hit-testing** follows what is drawn: `elementsFromPoint` honours the transform; the hairline test maps the pointer back to the file frame (`nearOutline`, `sizeOf`), so a stretched line is picked along its new path.
- **The centre clamp** of a move uses the centre as drawn (the file centre plus the resize shift).
- **API additions**: option `onResize({ id, w, h, dw, dh, phase })` (phase `start` | `end` | `cancel`), `size(id)` (`{ w, h, dw, dh }`, `dw = dh = 0` when not resized), `resizing` (the handle being dragged, or `null`). `reset()` and `dispose()` also put the resized elements back; `dragging` is the id for a move or a resize.

## Limitations (on purpose)
- Text and pictures are **stretched** with the element (a distorted glyph shape); PowerPoint re-flows text. That belongs to the step that commits a resize, which also needs the group-local conversion of 3.2.3-B for sizes. Nothing here is committed.
- No aspect lock (Shift), no resize from the centre (Alt), no snapping, no minimum other than 2 pt, no keyboard resize, no undo control in the UI (Escape cancels only an open drag).
- The rotation handle is still inert.
- A resized element whose drawing is "shrink on overflow" text is not re-fitted in the preview.

## Tests
`node tests/selection.mjs`, sections "resize preview (Phase 3.2.4-A)", "... in the real UI: mouse" and "... touch". See TESTING.md.
