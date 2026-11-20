# Phase 3.2.3-A: drag the selected element (temporary move preview)

Press the **already selected** element and drag: it follows the mouse, pen or finger. The move is a **temporary preview** and nothing else: no Scene Model change, no PPTX / XML edit, nothing saved, no resize, no rotation. Selection (3.2.1), the resize handles (3.2.2-A) and the rotation handle (3.2.2-B) still look and hit-test as before; they simply travel with the element.

## Behaviour
- **What starts a drag**: a press on the element that is selected *and visually on top at that point* (hairline reach included), followed by travel of more than the existing tap limit (mouse 4 px, pen 6 px, touch 10 px). Less travel is still a click / tap. The element goes the full distance from the press point (no lag at the start). A drag is never also a click, so it does not reselect or clear.
- **What does not**: a press on an unselected element, on empty space, or where a higher element covers the selected one. Those keep the 3.2.1 behaviour (a swipe or text drag selects nothing). Text in an unselected element can still be selected with the mouse; in the *selected* element a drag moves it, and a double click without movement still selects a word.
- **Where it can go**: anywhere, but the element's centre is kept inside the slide, so it can always be grabbed again.
- **How long it lasts**: until the slide is drawn again (another slide, leaving and re-entering the Rendered tab) or `reset()` / `dispose()` on the controller. The selection itself is restored on a redraw as before; the element is back at its file position.
- **Status line** (Rendered tab): "Selected: Shape “name” (id). Moved in the preview by 60 pt, 40 pt (not saved)", in slide points. The tab gains no buttons or controls.

## How (all in `js/render/selection.js`, plus the status line in `js/ui/rendered.js`)
- **One CSS `translate`** (the individual transform property) on the drawn element and on its outline, written through the CSSOM (CSP: no inline style attribute) as `calc(var(--u) * dx) calc(var(--u) * dy)`. `translate` is applied *outside* the element's own `transform: rotate()`, so it is a pure move in slide space. Rotation, flips (the renderer mirrors the drawing, not the frame) and group transforms (a group child's `world` frame is already flattened into slide space) are untouched, and the outline with its 8 handles and the rotation knob moves with the element for free.
- **Zoom accuracy**: the offset is kept in slide **points** and scaled by `--u`, so it is exact at every scale and stays on the same slide position when the slide is resized during or after the move. The pointer is converted to slide points (client px / px per point, measured on the slide as drawn *now*) on press and on every move, so the element follows the pointer exactly at 50 %, 100 %, 150 % and when the zoom or page scroll changes mid-drag.
- **Hit-testing follows the moved element**: `elementsFromPoint` honours `translate`; the hairline test (`nearOutline`) subtracts the preview offset, so a moved line is hit where it is drawn and no longer at its old place.
- **Capture**: the root takes pointer capture when the drag starts (moves and the release outside the slide or window still arrive); the text-selection start and native drag-and-drop are blocked only while a drag is pressed or running, and `user-select` is restored afterwards.
- **Touch**: the selected element (only) gets `touch-action: none`, so a finger drag on it moves it instead of scrolling the page; it also gets `cursor: move`. Both are removed again when it is no longer selected. The slide root and every unselected element keep `touch-action: auto`, so scrolling and pinch-zoom work as before.

## Cancellation (the element goes back to where it was *before this drag*, an earlier kept move stays)
`pointercancel`; losing pointer capture before the release; a second finger touching down; a right / middle button press; Escape; the window losing focus; `select(other)`, `clear()` or `dispose()` during the drag. A mouse move that reports the primary button up (a release nobody saw) finishes the drag at the last position instead of letting the element stick to the pointer. Only the `pointerup` of the dragging pointer keeps the move. Listeners on `document` / `window` exist only while a drag is open.

## API additions (`attachSelection`)
`offset(id)` (preview offset in slide points), `reset()` (put every moved element back), `dragging` (id being dragged or `null`), option `onMove({ id, dx, dy, phase })` with phase `start` | `end` | `cancel`. `onChange` is not fired by a drag.

## For the step that commits a move (not done here)
The preview offset is in **slide space**. A committed move must convert it to the parent group's local coordinates (undo the group's rotation, flip and scale), and must go through the Scene Model; this phase deliberately does neither.

## Limitations
- Touch drags need the finger to land on a painted pixel of the selected element. A finger beside a hairline (inside its reach but outside its drawn pixels) hits the slide, which scrolls, and the browser cancels the gesture (handled: nothing moves). Mouse and pen have the full reach.
- A selected element that covers most of the slide takes over finger scrolling over its area; start the scroll outside it or tap empty space first.
- No snapping, no guides, no keyboard move, no undo control (Escape cancels only an open drag). The move is lost when the slide is drawn again.
- Pinch while dragging cancels the drag; it does not zoom the slide in the same gesture.
