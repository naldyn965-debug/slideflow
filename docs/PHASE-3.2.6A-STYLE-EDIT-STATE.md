# Phase 3.2.6-A: style edit-state (data only)

The first stage of style editing: the temporary edit-state can now keep **completed style edits** next to the moves, resizes and text edits. This stage is **data only**. It follows the design of the 3.2.6 audit. Nothing is drawn differently, no control exists, nothing is written to the XML and nothing is exported: `render.js`, `selection.js`, `text-edit.js`, `rendered.js`, the resolver and the Scene Model are not touched (a test checks that none of them mentions the new API). Phase 3.2.6-B (renderer, controller, UI) is not started.

Only `js/model/edit-state.js` changed in the code.

## The record
One record per element id, the **total** style change since the file, holding **only what differs from the file's resolved value**:

```
{ id, part, slideIndex,
  text?: { color?, size?, font?, b?, i? },   // applied to EVERY run of the shape
  fill?: { hex },                            // a SOLID fill
  line?: { hex },                            // colour of an outline that exists and is solid
  was:   { text?: { color, size, font, b, i }, fill?: { type, hex, alpha }, line?: { hex, alpha } } }
```
- `color` / `hex`: `'RRGGBB'`, stored upper case (`#` and lower case are accepted on input). A colour edit means alpha 1, which is what PowerPoint writes; a hyperlink run keeps the theme's hyperlink colour.
- `size`: **nominal** points (what `<a:rPr sz>` would hold), 1 to 4000, two decimals. The resolved `run.sz` of a shape with a stored `normAutofit fontScale` is already multiplied by that scale, so the comparison with the file divides it out (tolerance 0.1 pt for the resolver's own rounding) and `effectiveText` multiplies it in again.
- `font`: a typeface name (trimmed, at most 100 characters; no quotes, `<`, `>`, `&`, backslash or control characters; no theme reference such as `+mj-lt`). Compared with the file's face case-insensitively. Which script face (`latin` / `cs`) it will map to is **not decided here**.
- `b`, `i`: booleans.
- `was` is what the file has for exactly what the record holds, so a later reader can tell a stale record from a current one. For text properties it is the value every affected run has, or `null` when the runs differ (or the face is the theme's).
- A value equal to the file's is not stored (for a text property: equal for **every** affected run; when the runs differ, it is kept, because it changes something). When nothing differs the record is removed.

## API (`createEditState()`)
- `commitStyle(scene, id, patch)` with `patch = { text?, fill?, line? }`; inside `text`: `color`, `size`, `font`, `b`, `i`. A value sets, **`null` reverts** (a single property, or a whole group: `text: null`, `fill: null`, `line: null`), an absent key or `undefined` leaves it, a patch merges into the earlier record. Returns `{ ok: true, edit, changed }` (`edit` = a copy of the record, `null` when none is left) or `{ ok: false, reason }` with **nothing changed** (the whole patch is checked before anything is applied).
- Reasons: `not-styleable` (not a plain slide shape or connector line: pictures, tables, groups, alternate content, layout / master elements, unknown ids), `invalid` (malformed patch or value, unknown key, empty patch), `no-text` (text property for an element with no run to hold it), `no-color-target` (text colour for text that is only hyperlinks), `no-fill-target` (fill for a connector line), `picture-fill` (fill for a shape whose fill is a picture), `no-outline` (line colour for an element whose outline is not solid or not there). A revert (`null`) is always allowed for a styleable element. `not-styleable` is reported before `invalid`.
- `getStyle(id)`, `hasStyle(id)`, `stylesForSlide(part)`, `listStyles()`: copies; `listStyles` in the order the elements were first styled.
- `discardStyle(id)`: removes the style only.
- `effectiveStyle(e)` (and the module function `effectiveStyle(e, rec)`): the element's `{ fill, line, shadow, … }` with the edited colours applied, as new objects (a gradient replaced by a solid colour has no stops; the line keeps its width, dash and joins). Without a record it **is** `e.style`.
- `effectiveStyledText(e)` (and `effectiveText(e, textRec, styleRec)`, the new third argument is optional): the text with the text record **and** the text style applied to every run and to each paragraph's default run (so text typed into an empty paragraph has the style). Without a style record the result is exactly what it was. **`effectiveText(e)` of the state is unchanged: text edit only.**
- Existing members now cover styles too: `has(id)`, `discard(id)` (move, resize, text **and** style of that element), `clear()`, `size` (elements holding any edit, counted once), `version` (+1 per real change; none for a no-op or a refusal).

Not changed: `commit`, `commitResize`, `commitText`, `get*`, `frame`, `forSlide`, `list*`, `discardText`, `effectiveText(e)`. A style commit leaves every move, resize and text record, the frame and every list exactly as they were (tested), and later moves, resizes and text edits leave the style record alone.

## What the tests cover
`tests/selection.mjs`, section "Phase 3.2.6-A" (65 checks incl. the static guards), with the new fixture `tests/fixtures/p326a-styles.pptx` (`python3 tests/make_fixture_p326a.py`) and fabricated variants of its elements (picture fill, only-hyperlink text, resolver-rounded scaled size, no text, layout origin): every property set / merge / revert / equal-to-file / re-commit, mixed runs, hyperlinks, stored `fontScale`, the empty paragraph's default run and a typed `fill` run, every refusal reason and 50 malformed inputs (nothing changes, atomic patch), the limits, combined edits with move + resize + text, `discardStyle` / `discardText` / `discard` / `clear` / `size` / `version`, copies, effective style and text, the Scene Model byte-identical afterwards, and static guards (edit-state.js stays DOM-free and writes no XML; the renderer, controller, text editor, UI and resolver do not use the API).

## Limitations (exact)
- Data only: nothing applies a style record yet (no drawing, no controller call, no `reset()` integration, no UI, no undo), and nothing writes XML or exports. The controller's `reset()` does not know styles yet (stage B); `createEditState().clear()` / `discard()` do.
- Element level only: a text property applies to **every** run of the shape (mixed formatting is flattened, as PowerPoint does for a selected shape). No per-run or per-selection styling.
- Text colour: hex only, the theme link of the file's colour is not kept, alpha becomes 1; hyperlink runs are excluded from colour edits.
- Fill: a solid colour only (cannot be removed or made none; gradients and patterns are replaced; a picture fill is refused). Line: colour only, only where a solid outline exists; no width, dash, or adding an outline.
- Size: absolute points. For shapes with a stored `fontScale` the stored scale / `lnSpcReduction` are not recomputed (a renderer drawing an edited shape from full size, as in 3.2.5-C, is stage B's job).
- Font: any acceptable name is recorded; whether it can be drawn (only the self-hosted families are) and which script face it maps to are not checked here.
- Support notes of the scene (e.g. "gradient approximated") are not recomputed for a replaced fill.
- Not styleable: pictures, tables, groups, alternate content, layout and master elements.
