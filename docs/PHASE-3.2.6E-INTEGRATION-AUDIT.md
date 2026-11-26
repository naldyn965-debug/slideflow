# Phase 3.2.6-E: integration audit

Last stage of 3.2.6. It audits the whole path of style editing as one flow (upload, select, text edit, style, move, resize, navigation, return, tab change, reset) and the lifecycle around it (dialog close, reopen, a second deck). **Result: no verified bug; no code under `js/` changed.** The phase adds regression tests (`tests/selection.mjs`, section "Phase 3.2.6-E") and this note.

## What was checked
- **The flow** (real mouse and keyboard): text and style survive every step in any order; the outline stays on the element after move and resize; slide changes leave the bar idle at once and nothing selected; returning shows the kept text, style, position and size; a tab change and back keeps the slide, the selection and the bar in sync.
- **Reset style** removes only the style of the selected element: the edited text and the move and resize stay, other elements and slides are untouched. It is per element, as documented in 3.2.6-D (there is still no UI that undoes a move or a resize).
- **Typed but not committed**: a value typed in the size or font box and then a click on another element, slide or tab is applied to the element it was typed for (the box commits on blur, before the selection or slide changes).
- **Same-tick races**: a style change and a slide change together leave a single slide, no outline, no editor, an idle bar and a clean status line; the style is kept for the right element.
- **Lifecycle**: nothing is written to localStorage or IndexedDB; after close (even during text editing, after a drag, or with a typed value) no listener stays on `document` or `window`; reopening the same template or opening another gives the file's values only.
- **Hand probes without a check** (all correct): autofit text, flipped shape with shadow, shape in a group, narrow viewport, eight quick arrow-key slide changes.

## Observations that are not bugs (left as they are)
- After a text edit that empties or replaces text, the empty runs keep their formatting (by design since 3.2.5-B), so the bar can still show "Mixed" for a shape whose visible text has one style. Reset style and the controls behave consistently with the stored runs.
- The bar's status line keeps "Style changed in the preview" after a later move or resize, until the selection or slide changes.
- Alive mutations and why they are judged equivalent or defensive are listed in TESTING.md, section 3.2.6-E.

## Tests
selection 1208 (1181 + 27), render 675, e2e 108, analysis 206 (31 FOOD checks SKIPPED without the deck); all pass.
