# Phase 3.2.5-B: commit text edits to the in-memory edit-state

Text typed in the 3.2.5-A edit mode is now **kept** in the same temporary edit-state as moves and resizes, and drawn again on every redraw, slide change and tab change for as long as the deck is open. **Still no XML write, no export, nothing saved**: the edit-state sits next to the Scene Model (never in it) and dies with the opened deck.

## Which exits keep the text
- **Kept (committed):** a press outside (empty space or another element), selecting another element, `clear()`, `endEdit()`, the tab being hidden, a slide change / redraw / closing (`dispose()`), focus leaving the text.
- **Cancelled (the original text, or the text kept before, stays):** Escape, a structure violation (`end:structure`), `reset()`.
- `onEdit({ phase: 'end', reason, committed, changed, skipped })` reports it. Status line: "Esc cancels; a click outside keeps the change in memory (nothing is saved)"; after a kept change "Text changed in the preview (kept while this deck is open; nothing is saved)".

## Data structure (`js/model/edit-state.js`, third map, one record per element id = the TOTAL since the file)
```
{ id, part, slideIndex,
  paras: [ { p, runs: [ { r, text, was } ], fill?: { text } } ] }
```
`p` = index into `e.text.paras` (the shape's `<a:p>` order), `r` = index into `e.text.paras[p].runs` (line breaks count), `text` = the run's new text, `was` = its text in the file (a record is applied only where the file text is still `was`), `fill` = text typed into an **empty** paragraph. Only changed runs are stored; a run emptied entirely stays (`text: ""`); committing the file's own text removes the record. API: `commitText(scene, id, draft)` (`{ ok, edit, changed }` or `{ ok: false, reason: 'not-editable' | 'invalid' }`), `getText`, `hasText`, `textsForSlide`, `listTexts`, `discardText`, `effectiveText(e)`; `has`, `size`, `version`, `discard`, `clear` cover texts as well. `effectiveText` is pure: it returns NEW paragraph / run objects with only `text` replaced; `e.text` is never written.

## Formatting, RTL, bullets, line breaks, empty paragraphs
- **Formatting:** only `text` of a run changes. Size, font, colour, link, bold/italic, `rtl` flag, spacing are the run's own, byte for byte (tested through computed styles).
- **RTL / bidi:** paragraph `dir`, alignment, indent stay; strings are stored verbatim, so RLM, LRM, NBSP and every other character survive in the record and on the slide.
- **Bullets / levels / numbering:** paragraph structure never changes, bullets are not text; labels are identical after a commit.
- **Line breaks (`<a:br>`):** stay where they are and cannot be removed (Backspace / Delete next to one, a selection across one, and select-all are blocked; a break added or removed behind the rules' back is a structure violation). Each line's text goes to its own run; the break is never in the record.
- **Empty paragraphs:** stay (count, direction, height); text typed into one becomes a `fill`, drawn with the paragraph's **default run**; delete it again and the record is gone.
- **Move / resize:** three independent records per element id. A moved shape can be edited (the move is kept); a text-edited shape can be moved and resized (the text is drawn stretched); editing a **resized** shape is still refused (3.2.5-A). `reset()` forgets the slide's texts, moves and resizes, DOM byte-identical to the file.

## Resolver additions (`js/model/resolve.js`, additive; nothing in the renderer reads them)
- `run.src = { ci, s }`: `ci` = index of the run's `<a:r>` / `<a:fld>` among the children of its `<a:p>`, `s` = the script segment of its text (one `<a:r>` can resolve to several runs). **For a future export**: a later phase can write a changed run's text back to `<a:t>` by concatenating the segments of one `ci`.
- `para.defRun`: for a paragraph with no text run, a run with empty text formatted from its `<a:endParaRPr>` and the inheritance chain (`src.def: true`, `src.ci` = the endParaRPr's index or -1). Run building moved into one helper (`buildRun`) used for real runs and the default run, so the two cannot drift.
- `fonts.js` loads the default run's font too.

## How it is drawn
`render.js` handle gained `setText(id, text, refit)`: it builds the text box with the same builder as the first drawing and puts it where the old one was (same rectangle); the controller calls it after a commit and, on attach, for every committed text of the slide (like moves and resizes), then fits once. A redraw and a commit produce the same DOM (tested byte for byte).

## How a draft is read (`text-edit.js` `collect()`)
Original span ↔ run by **identity**. Text not in an original span (a span the browser made, a bare text node) goes to the original run before it, or at the start to the first run still there; a run that is gone is an empty run. When **all** the text of a run is deleted and others of the paragraph remain, the original span is put back in its place with a zero-width placeholder and the caret goes into it, so typing lands in that run with its formatting (Chromium would otherwise write a styled clone span that the page's CSP refuses); the placeholder is not text. Typing into an emptied paragraph or over all of a paragraph's text uses the first run (3.2.5-A rule).

## Limitations (on purpose, or known)
- **Script / font:** a run keeps its own formatting. Arabic typed into a Latin run keeps that run's Latin font stack and has no `lang`; Latin typed into an Arabic run keeps the Arabic run's. The text is not re-segmented by script, so PowerPoint's complex-script font is not applied to it; the browser still orders the text by the bidi rules. Fixing it needs the resolver to expose a style per script.
- ~~Text typed into a paragraph that has only line breaks is not kept (`skipped`).~~ **Corrected in 3.2.5-C:** it was worse (the session ended with `structure` and the whole draft was lost). Such paragraphs are now non-editable while the mode lasts; see [PHASE-3.2.5C-AUDIT-FIXES.md](PHASE-3.2.5C-AUDIT-FIXES.md).
- A paragraph's `defRun` uses the paragraph's direction for its script (an RTL empty paragraph gets the RTL font stack).
- Autofit: a shrink-autofit shape is re-fitted when a text commit redraws it (not live while typing); the file's stored `fontScale` is not updated (no export). Shapes with a stored `fontScale` were not re-fitted at all in 3.2.5-B; fixed in 3.2.5-C.
- Table cells, fields, vertical text, columns, resized shapes cannot be edited (3.2.5-A rules); no formatting, no paragraph insertion.
- The analysis / design panels read the XML and do not see edited text (on purpose; the Rendered tab's own details are refreshed after a commit since 3.2.5-C).
- Native undo is not available for the steps the editor does itself (typing into an emptied paragraph, put-back runs).
- **3.2.5-C:** an emptied paragraph (a run with text `""`) used to make the next edit of the shape impossible (the renderer's placeholder span was counted as a run); fixed, see [PHASE-3.2.5C-AUDIT-FIXES.md](PHASE-3.2.5C-AUDIT-FIXES.md).
