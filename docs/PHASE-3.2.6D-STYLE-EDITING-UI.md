# Phase 3.2.6-D: style editing UI

Fourth stage of style editing: controls in the Rendered tab that change the style of the selected element, through the controller's `style(id, patch)` (3.2.6-C). **No** change to the controller, the renderer, the edit-state, the Scene Model, the XML, export or undo/redo. Nothing is saved: a style change is kept while the deck is open, exactly like moves, resizes and text edits.

Files: `js/ui/style-bar.js` (new), `js/ui/rendered.js` (wires it), `css/styles.css` (static rules, the CSP forbids inline styles). Three old "the tab has no inputs" guards were narrowed to "no inputs outside the style bar" (`tests/render.mjs` once, `tests/selection.mjs` twice).

## The style bar
A toolbar (`role="toolbar"`) under the slide, always present (disabled with nothing selected, so the layout does not jump; the status line under it reserves two lines for the same reason):

| Control | What it does |
|---|---|
| Text colour (colour box) | colour of every non-link run of the element's text |
| Size (number box, 1 to 4000 pt, step 0.5) | font size in points (the nominal size, as in the file); Enter or leaving the box applies it; an empty box goes back to the file's size; out of range is refused with a message and the box goes back |
| Font (text box with a list) | font family; the list offers the 12 families the renderer can draw with a self-hosted equivalent, any other name is accepted; empty goes back to the file's font; a name the edit-state refuses (quotes, `<`, `&`...) is reported ("not valid") and the box goes back |
| **B** / *I* (toggle buttons, `aria-pressed`) | bold / italic for all runs; "mixed" runs show `aria-pressed="mixed"`, a click makes all bold / italic |
| Fill (colour box) | solid fill (replaces a gradient or "no fill"; the box is dashed when the fill is not a solid colour) |
| Outline (colour box) | colour of an existing solid outline |
| Reset style | `style(id, { text: null, fill: null, line: null })`: takes the whole style off; moves, resizes and text edits stay |

- **What is shown** is read from the edit-state (`effectiveStyle`, `effectiveStyledText`, `hasStyle`), never from the slide's DOM: the file's values until a change, then the changed ones. Runs that differ show "Mixed" (empty size, dashed colour, mixed button) instead of one run's value.
- **What is enabled**: text controls need drawn text (a paragraph with runs); the text colour also needs a run that is not a link; Fill needs a shape that does not have a picture fill (a connector line has no fill); Outline needs a solid outline; Reset needs a style change on the element. A picture, table, group, master/layout element or hidden element disables everything and the status line says so.
- **Every change goes through `controller.style()`**: the bar never calls the edit-state's commit / discard methods or the renderer (a test scans the code). After a change it shows the controller's answer ("Style changed in the preview (kept while this deck is open; nothing is saved)." or the reason for a refusal) and the slide's count line and details are read again (a style can add or remove a fit note).
- **Selection, text editing, move, resize** are untouched: the bar is outside the slide, a press on it does not clear the selection; style() ends text edit mode first and keeps the typed text (3.2.6-C); the drawn move/resize/RTL direction and language are unaffected (tested with real mouse input). The colour boxes apply while the picker is used (each `input` event is one change; newer changes win over older drawings that finish late, 3.2.6-C); size and font apply on Enter / leaving the box.
- **Keys**: keys typed in the bar stay in the bar (the preview dialog turns arrow keys into slide changes); Escape in the size or font box puts the value back and does not close the dialog.
- **Another slide / tab and back**: the bar is idle (nothing selected) at once when a slide is left, and the styles are drawn again from the edit-state when the slide returns (the bar shows the changed values when the element is selected again).
- **RTL**: the bar is not mirrored (the rest of the app's chrome is not); an RTL element keeps its paragraph direction and language, and its new font keeps the Arabic fallbacks in the stack.

## Limitations (exact)
- Element level only: a change applies to all text of the element (no style for a selected part of the text). Solid colours only, picked with the browser's colour box (no alpha, no theme colours, no palette); no outline width / dash / arrow, no shadow, no gradient editing, no alignment or spacing. The outline control only recolours an outline that is already solid.
- Not saved, not exported, no undo / redo: reload or closing the preview loses it; Reset style is the only way back (per element). Support notes and the Fonts list are not recomputed for a new font (as in 3.2.6-B).
- A font name the machine does not have is accepted and drawn with the self-hosted equivalent when there is one, otherwise with the generic fallback; the box does not say which.
- The bar is always shown (about 120 px under the slide, wrapping on a narrow screen); on a touch screen the boxes are 40 px high but this was not tried on a device. English labels only, like the rest of the UI.
- The `controller.ready` hook of 3.2.6-C is not used by the panel: the panel's scene keeps its fit notes between redraws, so the details are already right after a styled slide is drawn again.
- Tested in Chromium only; no screen-reader test (names, roles and `aria-pressed` are checked in the DOM).
