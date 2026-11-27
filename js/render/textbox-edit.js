// Text editing of a NEW text box (Phase 3.2.7-A). The structure-keeping editor of text-edit.js (no Enter, one run per span, paragraphs fixed) is for the file's own text; a new box has no structure to keep
// (its paragraphs ARE what is typed), so it gets a plain <textarea> over its text instead: Enter makes a new paragraph, the browser's own caret, selection, IME, paste and undo, and one direction per box
// while typing (dir="auto"); when the mode ends, every line becomes a paragraph of its own and takes its own direction (js/model/textbox.js). The drawn text under the textarea is only hidden, never
// changed, and the formatting is copied to the textarea through the CSSOM (the page's CSP refuses inline style attributes). Nothing is written to the scene or the XML and nothing is saved.
//
// Same session shape as text-edit.js (the controller treats both alike): { inner, active, ended, canCommit, locked, text(), collect(), end(reason) }.
// Exits: Escape (the typed text is thrown away: the box keeps what it had), a press outside / focus leaving / any other reason (the typed text is the draft). onEnd(reason, draft) runs once, after the
// textarea is gone; `look` = { color, family, size, bold, italic } of the default run, used only while the box has no text of its own; draft = { free: true, texts: [one string per line] } or null for Escape.
export function startFreeEdit(box, { point = null, texts = [''], look = null, onEnd = null } = {}) {
  const wrap = box.querySelector(':scope > .sf-text'); const inner = wrap && wrap.querySelector(':scope > .sf-text-in'); if (!wrap || !inner) return null;
  const doc = box.ownerDocument;
  const ta = doc.createElement('textarea');
  ta.className = 'sf-text-free'; ta.value = texts.join('\n'); ta.rows = 1;
  for (const [a, v] of [['spellcheck', 'false'], ['autocapitalize', 'off'], ['autocorrect', 'off'], ['dir', 'auto'], ['aria-label', 'Text box text'], ['aria-multiline', 'true']]) ta.setAttribute(a, v);
  // the look of the text it covers, through the CSSOM
  const run = inner.querySelector('.sf-r:not([data-ph])') || inner.querySelector('.sf-r') || inner; const cs = doc.defaultView.getComputedStyle(run); const ws = doc.defaultView.getComputedStyle(wrap);
  for (const k of ['fontFamily', 'fontSize', 'fontWeight', 'fontStyle', 'color', 'lineHeight', 'letterSpacing']) ta.style[k] = cs[k];
  for (const k of ['paddingLeft', 'paddingRight', 'paddingTop', 'paddingBottom']) ta.style[k] = ws[k];
  if (look && !inner.querySelector('.sf-r:not([data-ph])')) { // an empty box has no run to copy: its default run's look is handed in
    if (look.color) ta.style.color = look.color; if (look.family) ta.style.fontFamily = look.family; if (look.size) ta.style.fontSize = `calc(var(--u) * ${look.size})`; if (look.bold) ta.style.fontWeight = '700'; if (look.italic) ta.style.fontStyle = 'italic';
  }
  const first = inner.querySelector('.sf-p'); if (first) ta.style.textAlign = doc.defaultView.getComputedStyle(first).textAlign;
  wrap.append(ta); inner.classList.add('sf-text-covered');
  let live = true; let ended = false;
  const read = () => ta.value.replace(/\r\n?/g, '\n').split('\n');
  const onKeyDown = (ev) => { if (ev.key !== 'Escape') ev.stopPropagation(); }; // the preview dialog turns the arrow keys into slide navigation
  const onEsc = (ev) => { if (ev.key === 'Escape') { ev.preventDefault(); ev.stopPropagation(); end('escape'); } }; // before the dialog sees it (it would close)
  const onFocusOut = () => { setTimeout(() => { if (live && doc.hasFocus() && doc.activeElement !== ta) end('blur'); }, 0); }; // focus left the text (not: the window lost focus)
  ta.addEventListener('keydown', onKeyDown); ta.addEventListener('focusout', onFocusOut); doc.addEventListener('keydown', onEsc, true);
  ta.focus({ preventScroll: true }); ta.setSelectionRange(ta.value.length, ta.value.length);
  void point; // (the caret goes to the end of the text: a textarea has no hit-testing of the drawn text under it)

  function end(reason = 'end') {
    if (!live) return; live = false;
    const draft = reason === 'escape' || reason === 'reset' ? null : { free: true, texts: read() };
    ta.removeEventListener('keydown', onKeyDown); ta.removeEventListener('focusout', onFocusOut); doc.removeEventListener('keydown', onEsc, true);
    ta.remove(); inner.classList.remove('sf-text-covered');
    ended = true; onEnd?.(reason, draft);
  }
  return {
    inner: ta,
    get active() { return live; },
    get ended() { return ended; },
    get canCommit() { return true; },
    get locked() { return 0; },
    collect() { return { free: true, texts: read() }; },
    text() { return read(); },
    end,
  };
}
