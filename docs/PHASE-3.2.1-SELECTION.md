# Phase 3.2.1: element selection

Selection only. No moving, resizing, text editing, toolbar, undo/redo or export; nothing is written, the OOXML is never touched.

## What it does
- Click or tap an element in the Rendered tab: it is selected by its stable Scene Model id (`ppt/slides/slide1.xml#3`, the same string as `scene.byId` keys and `data-id` on the drawn element) and outlined.
- Click or tap empty space: selection is cleared. Selecting another element replaces the selection.
- Slide change (filmstrip, keyboard, Slides tab): selection is cleared. Showing the Rendered tab again on the same slide keeps it.
- The status line under the slide names the selected element and its id (`aria-live`).

## Files
- `js/render/selection.js` (new): `attachSelection(handle, scene, { onChange })` returns `{ id, pick(x, y, pointerType), select(id), clear(), dispose() }`. No dependency on the store, Firebase or the XML.
- `js/ui/rendered.js`: attaches it after each draw, keeps `{ slide, id }`, status line.
- `css/render.css`: one static `.sf-sel` rule.
- Tests: `tests/selection.mjs`, `tests/make_fixture_p321.py`, `tests/fixtures/p321-selection.pptx`.

## Rules
- Selectable: every drawn element of the slide (shapes, pictures, lines, tables, labelled boxes of charts / SmartArt / OLE). Not selectable: groups (their children are drawn separately, the child is selected), hidden elements, the non-drawn branch of `mc:AlternateContent`, and master / layout elements. Template decoration is drawn but inert, as in PowerPoint's slide view; otherwise a full-slide template shape would make "click empty space" impossible. One predicate (`selectable()`) decides this.
- Hit-testing uses what the browser paints (`document.elementsFromPoint`: rotation, flips, clip paths and z-order for free). Counts: a shape's fill and outline, a picture inside its crop geometry, a text frame, a table, the box of an undrawn element. Does not count: empty corners of an ellipse's frame, the inside of an outline-only shape, a diagonal line's bounding box. The element visually on top wins.
- Hairlines (a 1 pt line is under 1 px wide) get a reach around the pointer measured against the real path: mouse 3 px, pen 6 px, finger 10 px. A hairline within reach wins only if it is drawn above whatever is exactly under the pointer.
- A selection is a click or tap: press and release of the same primary pointer within 4 px (mouse), 6 px (pen) or 10 px (touch), not cancelled. Drags, text selection, scrolling, pinch and non-primary buttons do not select. Nothing is prevented or captured, so text can still be selected and copied.
- The outline is one `<div class="sf-sel">` appended to the slide root only while something is selected. It copies the element's own frame style (rotation and scaling follow, also on resize), is a 2 px inside border with a white hairline, at least 6 px thick so a zero-height line shows, and has `pointer-events: none`. With nothing selected the drawn DOM is byte-identical to Phase 3.1.

## Limitations
- Groups cannot be selected as a whole yet (no drill-down model).
- Picture hits use the crop geometry, not the image's transparency.
- The outline is the element's frame (as in PowerPoint), not its shape; for a skewed element (non-uniform group scale under rotation) it follows the drawn, unskewed frame.
- Hairline reach uses a scratch `<svg>` that exists only during a pick.
- No keyboard selection and no Escape-to-clear; no multi-select; no hover highlight.
- Empty placeholders draw nothing, so they cannot be selected.
