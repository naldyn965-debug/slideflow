# Phase 3.2.6-C: the controller connects style records and renderer

Third stage of style editing. The selection controller (`js/render/selection.js`) now connects the edit-state's style records (3.2.6-A) to the renderer's `handle.setStyle` (3.2.6-B). **No** UI, XML, export or undo/redo. The controller asks the edit-state what a record is and asks the renderer to draw it; it decides neither (a test scans the code).

Code changes: `js/render/selection.js` (the integration) and one line of `js/render/render.js` (a stale-draw fix found by the tests, below).

## What the controller does (needs an `edits` option)
- **Drawn on attach.** A controller attached to a slide (a redraw, another slide and back, another tab and back) draws the committed style records of its slide from the edit-state, after the committed moves, resizes and texts. `controller.ready` resolves when they are drawn.
- **`await controller.style(id, patch)`** commits `patch` (`{ text?, fill?, line? }`, `null` reverts, see 3.2.6-A) with `edits.commitStyle` and draws the element's record. It resolves to the commit's own `{ ok, reason? }` and, for a commit that changed something, `{ edit, changed, drawn, stale }`. Extra reasons: `no-edit-state`, `not-styleable` (also for an element that is not drawn / selectable, e.g. hidden), `disposed`. A refusal or a no-op (`changed: false`) draws nothing; the selection is untouched.
- **Text edit mode ends first** (`onEdit` phase `end`, reason `style`; the draft is kept like on any exit but Escape) because the draw replaces the text box under the caret.
- **`reset()`** now forgets the style records of the slide with the other committed edits and puts the file's style back. Before, a record was dropped only for an element that was also moved (`edits.discard(id)`), and the slide was never redrawn.
- **`dispose()`** leaves drawn styles as they are (like drawn texts); the next controller draws them again from the edit-state.

## Stale draws
- Every draw of an element is numbered. The renderer drops an older draw that finishes after a newer one or after `reset()` (per-element call counter, 3.2.6-B); the controller reports it as `stale: true, drawn: false`.
- While a style is being drawn (a face may still be loading) text edit mode is **refused** (`onEdit` phase `refused`, reason `style-pending`): the draw would replace the text box the session holds. This also holds for the records drawn on attach, until `ready`.
- `style()` on a disposed controller is refused; a draw still running on a slide that was drawn again ends quietly on the old DOM.
- **Renderer fix (`render.js`, one line):** `setStyle` rebuilt the text from the text it had read when it *started*; a `setText` (a committed text edit) that arrived while it waited for a face was overwritten by the older text. It now draws the text as it is when the face has arrived. The controller's own guard makes this unreachable through the controller; it was reachable through the renderer API.

## Limitations (exact)
- Still nothing in the UI called `style()` at this stage; records came only from API calls. (Since 3.2.6-D the style bar of the Rendered tab calls it, see PHASE-3.2.6D-STYLE-EDITING-UI.md; `rendered.js` did not need `controller.ready`: its scene keeps the fit notes between redraws, so the details are right after a styled slide is drawn again.)
- A font that loads slowly shows the file's text style for a moment after a redraw (the style is drawn when the face has arrived); text edit mode cannot start meanwhile.
- The drawn style is not re-synchronised if something other than the controller changes the edit-state's styles while a slide is shown (e.g. direct `edits.commitStyle`): it shows at the next attach. Use `controller.style()`.
- A controller without an `edits` option cannot style (`no-edit-state`): there is no preview-only style.
- Everything else of 3.2.6-A / B stands: element level only, solid fill and an existing solid outline's colour only, support notes and the Fonts list are not recomputed, nothing is written to the XML or exported, no undo/redo.
