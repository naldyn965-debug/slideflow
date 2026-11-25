# Phase 3.2.5-A: inline text edit mode (preview only)

> **Update (Phase 3.2.5-B):** the draft is now kept in the temporary edit-state on every exit except Escape, a structure violation and reset(); see PHASE-3.2.5B-TEXT-COMMIT.md. Statements below that the original DOM is restored "on every exit" describe 3.2.5-A; in 3.2.5-B the DOM is still restored from the original first, and a kept draft is then drawn again from the edit-state.

Double-click (mouse) or double-tap (finger, pen) a text shape and its drawn text becomes editable in place. It is a **temporary draft in the DOM**: every way out puts the **original DOM back, byte for byte**. Nothing is written to the Scene Model, the XML or the edit-state, nothing is saved. Text persistence is a later phase (3.2.5-B).

## Behaviour
- **Enter:** a mouse double-click on a text shape (the word under the pointer stays selected, so typing replaces it); a pen / finger: two taps within 350 ms and 24 px on the same element. Two bare clicks never enter (a mouse enters by double-click only). `beginEdit()` is the same thing for code and tests.
- **Eligible:** a slide text shape (`sp`) with drawn text, **no committed resize** (a previewed or committed resize makes it refuse), **no field** (`a:fld`: slide number, date), **no vertical text**, **one column**. Everything else is refused with a reason (`not-text-shape`, `no-text`, `resized`, `field`, `vertical`, `columns`) and a message on the status line; the element just stays selected. Pictures, tables, charts, SmartArt, connectors, layout / master elements are never editable. A moved text shape is editable.
- **While editing:** the controller stands down for that element (no handles, no drag, no resize; a press on the shape, including its padding, keeps the caret). A press anywhere else, `select()`, `clear()`, `reset()`, `dispose()` (a redraw, another slide, closing), the tab being hidden, Escape, or focus leaving the text ends it. The window merely losing focus does not (the draft survives an alt-tab).
- **Keyboard:** Escape cancels (it does not close the preview dialog); every other key stays in the text (the dialog would take the arrow keys for slide navigation).
- **Status line:** "Editing the text of Shape “name” in the preview. Esc, or a click outside, puts the original text back; nothing is saved." A window resize during editing is applied when the mode ends, not under the caret (the text is not re-fitted while typing).

## What is kept fixed (the structure is the file's, only characters are a draft)
- **No paragraph split, no line break:** Enter and Shift+Enter are blocked. Backspace at the start of a paragraph, Delete at its end, a selection across paragraphs (typed over, deleted, cut) are blocked; so are formatting commands (bold, italic, underline), drop and drag of text, and rich paste (paste is plain text, line breaks become spaces).
- **Bullets** stay: they are non-editable, nothing may start before their end or touch them, and a caret that lands before a bullet (Home, arrow keys) is moved to just after it.
- **Empty paragraphs** stay (the zero-width placeholder is not text; typing into one fills only that paragraph).
- **Characters are never normalised:** NBSP, LRM, RLM and every bidi mark, already in the text or typed, stay exactly as they are. The paragraph keeps its `dir` (a paragraph is RTL or LTR as the file says, whatever script is typed), script spans keep their `lang`.
- Typing into an **emptied paragraph or over all of a paragraph's text** goes into the paragraph's original first run (it keeps that run's formatting). The browser would otherwise write a typing-style `style` into the attribute text, which the page's CSP refuses.
- **Paste and programmatic commands:** `execCommand` does not give the page a cancelable `beforeinput`, so paste applies the same range rules itself (refused across paragraphs or in front of a bullet; into an emptied paragraph it uses the first-run rule). Anything that still gets through (a paragraph inserted by a command, text before a bullet: a bullet must be the very first node) is caught by the structure check after the fact.
- If something still changes the structure (an IME, a browser quirk) the edit ends and the original is restored (`end:structure`).

## Restoring the original
The original nodes are never cloned (a clone carries inline `style` attributes through `setAttribute`, which the CSP refuses and which would silently drop the styles). They stay in place; on exit the recorded children, text, attributes and (through the CSSOM) styles of every original node are put back, and anything the browser added is dropped. The result is byte-identical outerHTML **and** identical computed styles.

## Files
New `js/render/text-edit.js` (eligibility + session). `js/render/selection.js`: `editing` guard, `beginEdit` / `endEdit` / `editText` / `editing`, `onEdit` callback, double-click and double-tap entry (the file never touches the text DOM). `js/ui/rendered.js`: status line, deferred refit, hide ends the mode. `css/render.css`: edit-mode rules. Not touched: `resolve.js`, `scene.js`, `edit-state.js`, any XML.

## Limitations (on purpose, or known)
- Preview only: the draft is lost on every exit.
- No keyboard entry (nothing is focusable yet); no paragraph or line-break insertion; no formatting; no editing of fields, vertical text, columns, resized shapes, tables, empty text shapes.
- A script-span boundary: typing lands in the neighbouring run, which keeps that run's font; re-segmentation by script is for 3.2.5-B.
- `text()` (the draft read-back) leaves out bullets, the empty-paragraph placeholder and the browser's placeholder `<br>`; a real `<a:br>` and the browser's placeholder are told apart only by "the only child".
- The browser's native undo is not available for the one step that types into an emptied / wholly-selected paragraph.
- IME composition is not cancelable: the structure check catches anything it does.
