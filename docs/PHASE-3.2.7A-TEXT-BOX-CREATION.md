# Phase 3.2.7-A: text box creation

The Rendered tab can make new text boxes on the current slide. Preview only: nothing is saved or exported, the boxes live while the deck is open (as moves, resizes, text and style edits do).

## How it fits
- **A box is a record** in the edit-state (`addTextBox`, `setAddedText`, `getAdded`, `addedFor`, `removeAdded`, `addedLook`): id `slide#xmlId`, frame in EMU, one `{ text, rtl }` per paragraph. It is its own "file": its text edits replace the record's paragraphs and never make a text record; its moves, resizes and styles are ordinary records over its id. `discard()` and the controller's `reset()` keep the box.
- **`js/model/textbox.js`** (pure, no imports) turns a record into a detached `<p:sp>` (a text box: `txBox`, rectangle, no fill, `normAutofit`, per-paragraph `rtl`/`algn`) created in the slide's own document and attached nowhere. `buildScene(parsed, i, { added })` walks it after the file's shapes (on top), flagged `added`; the resolver, renderer, selection, move, resize and style bar see an ordinary text shape. The PPTX XML and zip are never touched.
- **`js/render/textbox-edit.js`** is the editor of a new box: a `<textarea>` over the text (formatting copied through the CSSOM, CSP-safe), `dir="auto"`, Enter = new paragraph, Escape cancels, focus leaving or a press outside keeps the text. It has the session shape of `text-edit.js`, which the controller routes to for added ids (`beginEdit` / `afterEdit`); the structure-keeping editor of the file's own text is unchanged.
- **`js/ui/rendered.js`**: the button; the scene of a slide is built again when its boxes change; after Add the box is selected and its editor opens when the slide is drawn (retrying while the slide's styles are still being drawn); an edit's redraw waits until a press ends; an empty box is removed (status "Empty text box removed.").
- The renderer draws an empty new box (it is where the text is typed); the style bar treats a new box as having text.

## Behaviour
- Default frame 40% of the slide wide, about one line high, centred at 40% of the height; each further box on a slide is offset a little down and right.
- Empty (nothing but white space, line breaks, zero-width characters) means removed: Escape or a press outside with nothing typed, or all text deleted. Escape on a box with text restores its text.
- A paragraph's direction follows its first strong letter when the edit ends (right to left paragraphs are right aligned).

## Limits
- A resized box cannot be text-edited (the preview stretches its text, as for every shape): edit first, then resize.
- The box does not grow with its text: long text shrinks to fit (`normAutofit`).
- While typing there is one direction for the whole textarea; the caret goes to the end when the editor opens.
- Reset style keeps the box and its text; a box with text is removed only by deleting its text.
- No save, export, undo or deleting by key; no touch, IME, other browser or screen reader testing.

## Tests
selection 1251 (1208 + 43), render 675, e2e 108, analysis 206 (31 SKIPPED); mutation checks in TESTING.md (22 of 24 detected, 2 defensive).
