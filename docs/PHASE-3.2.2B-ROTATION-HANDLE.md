# Phase 3.2.2-B: rotation handle (visual only)

The selected element's handle layer (Phase 3.2.2-A) gets one more child: a round knob 24 px above the top-centre resize handle, joined to the frame's top edge by a 1 px stem. **Visual only**: no rotating, dragging, resizing or PPTX change. Like the resize handles it has `pointer-events: none`, no cursor and no listeners, so selection and touch are unchanged (a click on it goes to whatever is under it).

## How
- `js/render/selection.js`: `drawOutline()` appends `<span class="sf-rot">` to the `.sf-sel-h` layer (before the 8 `.sf-hdl`, so the top-centre handle covers the stem's end). It exists only while something is selected and goes with the outline on `clear()` / slide change / `dispose()`. It is not a `.sf-hdl`, so the "8 resize handles" count is unchanged.
- `css/render.css`: same declarations as `.sf-hdl` (one shared rule: 8 px, white, accent border, soft shadow), round, `left: 50%; top: calc(-1 * var(--sf-rd))` with `--sf-rd: 24px` (28 px on coarse pointers); the stem is `::before`.
- Alignment for free: the knob lives in the outline's rotated frame, so "above" means above in the element's own frame. It follows rotation (hand-computed 45 degrees in the tests), nested group transforms, window resize and slide zoom. The 24 px is in screen pixels, so zoom moves the knob with the frame but never stretches the distance.

## Limitations
- The slide clips its content: for an element whose top edge is within ~24 px of the slide's top, the knob (and part of its stem) is cut off.
- It is round while the resize handles are square (deliberate, to tell them apart); colours, border, size and shadow are identical.
- For a vertically flipped element the knob sits above the drawn frame's top edge, which for a flipV element in PowerPoint terms is the visual bottom (the renderer draws flipV as flipH + 180 degrees; the outline follows what is drawn).
- Not focusable or announced (aria-hidden); no keyboard path yet.
