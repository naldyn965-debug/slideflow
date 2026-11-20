# Phase 3.2.2-A: resize handles (visual only)

While an element is selected (Phase 3.2.1) its outline carries 8 handles: 4 corners and 4 edge centres (`nw n ne e se s sw w`). **Visual only**: no dragging, resizing, rotation or PPTX change. The handles have `pointer-events: none`, no listeners and no resize cursor, so selection, hit-testing and touch behave exactly as in 3.2.1 (a click on a handle goes to whatever is under it).

## How
- `js/render/selection.js`: `drawOutline()` appends one `.sf-sel-h` container with 8 `.sf-hdl` spans to the outline. They exist only while something is selected and go with the outline on `clear()` / slide change / `dispose()`.
- `css/render.css`: the container covers the outline's border box (`inset: -2px`), each handle sits at 0 / 50 % / 100 % of it and is centred with `translate(-50%, -50%)`.
- Accuracy for free: handles are children of the outline, which copies the element's own frame style. Rotation, nested group transforms, resizing the window and slide zoom (`--u`) all follow; the handles themselves rotate with the element. Handle size is in screen px (8 px, 10 px on coarse pointers; white, 1.5 px accent border, soft shadow), so zoom never resizes them.
- Very small frames: edge-centre handles that would pile onto the corner handles are hidden by container queries (frame < 32 px along their edge, or < 12 px across it). All 8 stay in the DOM. A hairline shows its 4 corners.

## Limitations
- Handles centre on the frame edge, and the slide root clips like the slide does: for an element touching the slide edge the outer half of the handles on that edge is cut off.
- Corner handles of a hairline or a tiny element overlap each other.
- Container queries need Chrome 105 / Safari 16 / Firefox 110; older browsers show all 8 handles on tiny frames.
- A skewed element (non-uniform group scale under rotation) gets handles on the drawn, unskewed frame.
- Handles are not focusable or announced (aria-hidden); there is no keyboard path yet.
