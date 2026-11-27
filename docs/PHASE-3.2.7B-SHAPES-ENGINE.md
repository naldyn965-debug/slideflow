# Phase 3.2.7-B: shapes engine

The Rendered tab can add a rectangle, a rounded rectangle, an ellipse, a line and an arrow to the current slide, and remove what it added. Preview only: nothing is saved or exported; the elements live while the deck is open.

## How it fits (the text box infrastructure of 3.2.7-A, generalised)
- **A record per element** in the edit-state, now with a `kind` (`textbox`, `rect`, `roundRect`, `ellipse`, `line`, `arrow`): `addElement(scene, kind)` (`addTextBox` is the `textbox` case), `getAdded`, `addedFor`, `removeAdded`, `isAdded`; `isAddedText(id)` tells the text boxes (the only ones with text of their own; `setAddedText` refuses a shape). Moves, resizes and styles are ordinary records over the element's id.
- **`js/model/textbox.js`** builds the detached element of a record: `shapeElement()` makes a `p:sp` (rect / roundRect / ellipse: theme accent 1 fill, darker accent 1 outline, 1 pt) or a `p:cxnSp` (line: `line`, arrow: `straightConnector1` with a triangle `tailEnd`; 2.25 pt in the text colour). Colours are theme colours, so a shape follows the deck. `shapeFrame()` gives the frame: 25% × 20% of the slide centred for a closed shape; 30% of the width, flat, at mid height for a line or an arrow; further elements offset down and right.
- **`scene.js`** walks the detached elements after the file's own (on top) and sets `e.added` to the kind. The renderer, selection, move, resize, style bar and resolver see ordinary shapes and connectors: the style bar offers fill + outline for a closed shape and the outline only for a line or arrow; `render.js` and `style-bar.js` treat only `added === 'textbox'` as text boxes.
- **`js/ui/rendered.js`**: six add buttons and Remove (enabled only for elements the preview added); a shape is selected right after it is made (no editor); Remove deletes the element with its records and redraws; the count line says "N new elements included".
- The PPTX XML and zip are never touched.

## Limits
- No text in a new shape yet (a double click says so); shapes cannot be removed by key, only with Remove; the file's own elements cannot be removed.
- Line width, dash, arrow style and shadow are not editable (the style system edits colours only); lines and arrows start horizontal and are lengthened or moved, not rotated or flipped.
- A resized closed shape is stretched like any other shape; no save, export, undo, layers or grouping.
- No touch, other-browser or screen-reader testing.

## Tests
selection 1297 (1251 + 46), render 675, e2e 108, analysis 206 (31 SKIPPED); mutation checks in TESTING.md (17 of 18 detected, 1 equivalent).
