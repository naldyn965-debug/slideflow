# Phase 3.2.5-C: the four fixes of the 3.2.5-B audit

An audit of the 3.2.5-B limitations (script / font handling, line-break-only paragraphs, IME and mobile composition, autofit after edits, the analysis panels) found two bugs that silently lost typed text and two small gaps that were worth closing before anything is built on the text commit. This phase fixes exactly those four. Nothing else changed: **no XML write, no export, no change to the Design tab, no mixed-script font change**, no new capability for the person (no typing into line-break-only paragraphs yet).

## 1. An emptied paragraph no longer blocks the shape (critical)
**Bug.** Delete all the text of a paragraph and commit. The redraw drew the run's empty span *and* the renderer's zero-width placeholder span, so on the next edit the span list had one span more than the paragraph had runs, the editor decided it could not map spans to runs (`canCommit = false`) and dropped the whole draft. From then on **every** edit of **any** paragraph of that shape ended with `committed: false`, no message and no error, until `reset()`.

**Fix.**
- `render.js`: the placeholder span carries `data-ph="1"`. It is the renderer's, not a run (no style attribute, so CSP is unchanged).
- `text-edit.js`: a placeholder span is never matched to a run (`slots`), its zero-width character is never text, and text typed into it belongs to the run before it (or is the paragraph's leading text, as before). `isEmptyPara` now means "no line break, and no run holds text" (an empty paragraph of the file, **or** one whose runs were all emptied), so typing into such a paragraph uses the original first run (same lang / size / font) as it does for an empty paragraph of the file.
- The record itself did not change: an emptied run is still `text: ""`.

## 2. Line-break-only paragraphs are not editable while the mode lasts (critical, rarer)
**Bug.** A paragraph that is only `<a:br/>` (or several) has no run to put typed text into. Typing there made the browser remove the placeholder break, which is a structure violation: the session ended with `end:structure`, `committed: false`, and **everything typed earlier in that shape in the same session was thrown away**, silently. The 3.2.5-B documentation said the text was "skipped"; it was worse.

**Fix.** On entering the mode, every paragraph that has line breaks and nothing else (a bullet aside) gets `contenteditable="false"`, like the bullets. Nothing typed, pasted or composed (IME) can reach it, no structure violation can come from it, the paragraph's breaks are untouched. `restore()` puts the attribute back (the existing mechanism). Also:
- `structural()` refuses a range inside such a paragraph (defense in depth, e.g. for paste).
- The caret of a fresh edit goes into the last paragraph that *can* be typed into.
- A shape whose paragraphs are **all** line-break-only is refused: reason `line-breaks-only`, "This text holds only line breaks: there is nothing to type into yet."
- `onEdit({ phase: 'start', locked })` says how many paragraphs were locked; the status line adds "A paragraph that holds only a line break cannot be typed into yet."

A paragraph with text **and** a line break (run, break, run) is unchanged: still editable.

## 3. Shrink autofit of an edited shape with a stored fontScale (necessary now)
**Bug.** `normAutofit` with a stored `fontScale` (PowerPoint writes it whenever it shrank the text; very common) was drawn at the file's scale and never fitted: added text overflowed the frame, shortened text stayed small.

**Fix.** `setText(id, text, refit, edited)` has a fourth argument, `edited` (the controller passes `edits.hasText(id)`; on attach every committed text is `true`; `reset()` leaves it `false`). For an edited shape with `autofit = shrink` and a stored scale, the text box gets `data-autofit="shrink"` and `data-fit-base="<stored scale>"`. `fitText()` then starts from the **full size** (`--fit = 1 / base`, the sizes the file had before PowerPoint shrank them) and shrinks in 5 % steps of the PowerPoint scale until the text fits: text that grew shrinks, text that got shorter grows back. A shape that has no committed edit, or whose text was typed back to the file's, is drawn exactly as before (the stored scale). Shapes without a stored scale behave as before.

## 4. No stale fit notes; the Rendered details follow a commit (necessary now)
**Bug.** `fitText()` wrote its note ("text overflows its frame … scaled to N%") into the cached scene element and never removed it: the support level stayed "approximated" after the text was back to normal, a note per different percentage piled up, and the Rendered tab's count line and details were only rebuilt on a redraw.

**Fix.**
- `render.js` `clearFitNote(e)`: before each fit, and when a text box is replaced (`setText`), the element's fit note is removed and its support level goes back to what the other notes (or the level before the fit note) say; `data-support` is resynchronised. A note of any other origin stays.
- `rendered.js`: one `refresh()` (fit, count line, details) is used for the first drawing and after every **committed and changed** text edit.

## Not changed (see Limitations)
Design tab and the stored analysis (they describe the file, not the preview edits); mixed-script fonts; export; typing into line-break-only paragraphs; the 3.2.5-A input rules.

## Limitations (known, left on purpose)
- **Script / font** (unchanged from 3.2.5-B, audited): the stored text is right, only the preview differs. Arabic typed into a Latin run is painted with the system fallback (the non-RTL font stack has no Arabic fallback); Latin typed into an Arabic run is painted with the Arabic stack (DejaVu: wider than Calibri / Carlito, so lines can wrap differently); the span keeps its `lang`. The fonts list of the Rendered tab reads the file's runs, not the edited text. Later fix: expose the Latin and the complex-script face per run and build the stack from both plus the Arabic fallbacks.
- **Line-break-only paragraphs** cannot be typed into at all yet (they used to lose the session). Allowing it needs a `fill`-like record for such a paragraph and a position for the text relative to the break.
- **Autofit:** the fit is computed when the mode ends (not live while typing). `lnSpcReduction` stored in the file stays as drawn when an edited shape grows back to full size (PowerPoint pairs it with `fontScale`). The file's stored `fontScale` / `lnSpcReduction` are not written anywhere (no export): a later export has to recompute or drop them.
- **IME / mobile:** simulated compositions in desktop Chromium keep the mode and the text (mid-run, replacing a whole paragraph, into an emptied paragraph, over a put-back run, interrupted by a press outside). Composing into an emptied paragraph still makes the browser write a styled clone span, which logs the CSP "Applying inline style" console error (the text is right). **Not tested on real devices or other engines:** Gboard / iOS keyboards, Firefox, WebKit; the editor rebuilds spans and moves the caret during composition. The status line still says "Esc cancels", which a touch keyboard cannot do; no keyboard-aware scrolling (focus uses `preventScroll`).
- **Panels:** the Design tab, the stored analysis and the filmstrip titles read the file, not the edits (on purpose until there is an export to re-analyse). The Rendered tab's details are rebuilt after a text commit, not after a move, a resize or `reset()`, and not when a resize of the window changes a fit while no commit happens.
- Everything else of 3.2.5-B (tables, fields, vertical text, columns and resized shapes cannot be edited; no formatting, no paragraph insertion; native undo does not cover the editor's own steps).
