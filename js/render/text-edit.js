// Inline text editing (Phase 3.2.5-A edit mode, Phase 3.2.5-B keeping the text in the temporary edit-state).
//
// Entering edit mode makes the drawn text of ONE selected text shape editable in place (`contenteditable="plaintext-only"` on its `.sf-text-in`), so a person can type, delete and move the caret. It is a
// temporary draft in the DOM: every way out (Escape, a press outside, another selection, another slide, the tab being hidden, dispose) puts the ORIGINAL children back, byte for byte, from a
// record taken when the mode began. On every exit except Escape, a structure violation and reset() the draft is also READ (collect(): the changed runs' text, by original span) and handed to the caller, which
// keeps it in the edit-state and draws the element's text again from there; the Scene Model and the XML are never written and nothing is saved. This module only reads the scene element (to
// decide eligibility, including a look at the XML for fields) and writes to the one text node it was given.
//
// What is kept fixed while editing (the structure of the text is the file's, only the characters are a draft):
//   - paragraphs: Enter and Shift+Enter are blocked (no paragraph split, no new line break), and so is every edit whose range starts or ends in another paragraph (Backspace at the start of one, Delete at
//     the end of one, a selection across paragraphs, select-all and type). Paste is plain text with line breaks turned into spaces; drop is refused; formatting commands (bold, italic, ...) are refused.
//   - bullets: they are drawn inside the paragraph; they are made non-editable and no edit may start before their end or touch them.
//   - empty paragraphs: kept (the zero-width placeholder span, marked data-ph by the renderer, is not a run and is left alone, also in a paragraph whose runs were all emptied: Phase 3.2.5-C); text() leaves it out.
//   - line-break-only paragraphs (breaks and nothing else): made non-editable (contenteditable="false") while the mode lasts, so nothing typed, pasted or composed can reach them (3.2.5-C); a shape of nothing else is refused (line-breaks-only).
//   - characters: nothing is normalised. NBSP, the bidi marks (LRM, RLM, ...) and every other character typed or already there stay exactly as they are; the paragraphs keep their `dir`, the script spans their `lang`.
//   - if anything still changes the structure (an IME, a browser quirk) the edit ends and the original is restored: the draft is never allowed to become a different text structure.
// While editing: Escape cancels (restores), arrow keys and every other key stay inside the text (the preview dialog would otherwise take the arrows for slide navigation), a press on the shape's padding keeps
// the caret, and focus leaving the text (not the window losing focus) ends the mode.

const ZWSP = '\u200b'; // the placeholder the renderer puts in an empty paragraph
const ALLOWED = new Set([
  'insertText', 'insertCompositionText', 'insertFromComposition', 'insertReplacementText',
  'deleteContentBackward', 'deleteContentForward', 'deleteWordBackward', 'deleteWordForward', 'deleteSoftLineBackward', 'deleteSoftLineForward', 'deleteHardLineBackward', 'deleteHardLineForward',
  'deleteByCut', 'deleteCompositionText', 'historyUndo', 'historyRedo',
]);
/** Exits that throw the draft away (the original text stays): Escape, a structure violation, reset(). Every other exit keeps the draft (see onEnd). */
const CANCEL = new Set(['escape', 'structure', 'reset']);
const ATTRS = ['contenteditable', 'spellcheck', 'autocapitalize', 'autocorrect', 'role', 'aria-multiline'];

/**
 * Can this scene element be edited inline? `node` = its drawn `.sf-el`, `resized` = it has a (committed or previewed) resize.
 * Returns { ok: true } or { ok: false, reason, message } (message is for the status line).
 * Eligible: a text shape (`sp`) of the SLIDE (not a layout or master element) that has drawn text, no resize, no field, no vertical text, a single column.
 */
export function textEditability(e, { node = null, resized = false } = {}) {
  const no = (reason, message) => ({ ok: false, reason, message });
  if (!e || e.origin !== 'slide' || e.tag !== 'sp' || e.kind !== 'shape') return no('not-text-shape', 'Only text shapes on the slide can be edited.');
  if (!node || !node.querySelector(':scope > .sf-text > .sf-text-in') || !e.text || !e.text.paras || !e.text.paras.length) return no('no-text', 'This shape has no drawn text to edit.');
  if (e.text.paras.every((p) => p.runs.length && p.runs.every((r) => r.br))) return no('line-breaks-only', 'This text holds only line breaks: there is nothing to type into yet.'); // (3.2.5-C)
  if (resized) return no('resized', 'A resized text shape cannot be edited yet: its text is stretched.');
  const b = e.text.body || {};
  if (b.vert) return no('vertical', 'Vertical text cannot be edited yet.');
  if (b.cols > 1) return no('columns', 'Text in several columns cannot be edited yet.');
  if (e.el && e.el.getElementsByTagNameNS('*', 'fld').length) return no('field', 'Text with fields (slide number, date, ...) cannot be edited yet.');
  return { ok: true };
}

const paraOf = (n) => (n ? (n.nodeType === 1 ? n : n.parentElement)?.closest?.('.sf-p') || null : null);
const bulletOf = (p) => p.querySelector(':scope > .sf-bu');
/** The renderer's placeholder span (Phase 3.2.5-C): the zero-width run it puts into a paragraph that shows no text. It is not a run of the file and never text. */
const isPh = (n) => n.nodeType === 1 && n.getAttribute('data-ph') === '1';
/** An empty paragraph as the renderer draws it: no line break, only runs (bullet apart) that hold no text, or the zero-width placeholder: an empty paragraph of the file, or one whose runs were all emptied by an edit. */
function isEmptyPara(p) {
  const kids = [...p.children].filter((c) => !c.classList.contains('sf-bu'));
  if (!kids.length || !kids.every((c) => c.classList.contains('sf-r'))) return false;
  const t = kids.map((c) => c.textContent).join('');
  return t === '' || t === ZWSP;
}

/**
 * Start editing the text of `box` (a drawn `.sf-el`). `point` = { x, y } client coordinates where the caret goes (end of the text when absent or not on the text).
 * `onEnd(reason, draft)` is called once when the mode ends, AFTER the original DOM is back; `draft` is the text change to keep (see collect) or null when the exit cancels (Escape, a structure violation, reset)
 * or nothing can be kept. `paras` = the paragraphs the DOM was drawn from (the committed text applied) and `orig` = the file's resolved paragraphs: without them nothing can be committed (the preview of 3.2.5-A).
 * Returns a session { inner, active, text(), collect(), end(reason) } or null when the element has no text box.
 */
export function startTextEdit(box, { point = null, onEnd = null, paras = null, orig = null } = {}) {
  const inner = box.querySelector(':scope > .sf-text > .sf-text-in'); if (!inner) return null;
  const doc = box.ownerDocument;
  // The original, restored on every exit. NOT cloned: a clone would carry the inline `style` attributes through setAttribute, which the page's CSP refuses (the styles would be lost). The original nodes stay
  // where they are while editing; what is recorded is enough to put every one of them back (children, text, attributes; a style only through the CSSOM).
  const texts = new Map(); const els = new Map(); const innerKids = [...inner.childNodes];
  (function record(n) { for (const c of n.childNodes) { if (c.nodeType === 3) texts.set(c, c.data); else if (c.nodeType === 1) { els.set(c, { kids: [...c.childNodes], attrs: [...c.attributes].map((a) => [a.name, a.value]), css: c.style ? c.style.cssText : '' }); record(c); } } })(inner);
  const sameList = (a, b) => a.length === b.length && a.every((x, i) => x === b[i]);
  function restore() {
    if (!sameList([...inner.childNodes], innerKids)) inner.replaceChildren(...innerKids);
    for (const [n, d] of texts) if (n.data !== d) n.data = d;
    for (const [el, r] of els) {
      if (!sameList([...el.childNodes], r.kids)) el.replaceChildren(...r.kids);
      const now = [...el.attributes].map((a) => [a.name, a.value]);
      if (JSON.stringify(now) !== JSON.stringify(r.attrs)) { // (the bullets' contenteditable; a style the browser's editing code wrote into the attribute text): remove what was added, put back what differs, leave the rest where it is
        const was = new Map(r.attrs);
        for (const [n] of now) if (!was.has(n)) el.removeAttribute(n);
        for (const [n, v] of r.attrs) if (n !== 'style' && el.getAttribute(n) !== v) el.setAttribute(n, v);
        if (el.style && was.has('style') && el.getAttribute('style') !== was.get('style')) el.style.cssText = r.css; // (through the CSSOM: setAttribute('style') is what the CSP refuses; this rewrites the attribute text from the declarations)
      }
    }
  }
  const prior = ATTRS.map((a) => [a, inner.getAttribute(a)]);
  const paras0 = [...inner.children]; const count0 = paras0.length;
  const bullets0 = paras0.map((p) => (bulletOf(p) ? 1 : 0)); const empty0 = paras0.map(isEmptyPara); const dirs0 = paras0.map((p) => p.dir); const firstRun = paras0.map((p) => [...p.children].find((c) => c.classList.contains('sf-r')) || null);
  const brs0 = paras0.map((p) => p.querySelectorAll('br').length); // real line breaks (<a:br>): they stay where they are
  const onlyBr = (p) => [...p.childNodes].every((n) => n.nodeType === 1 && (n.classList.contains('sf-bu') || n.tagName === 'BR'));
  const brOnly = paras0.map((p, i) => brs0[i] > 0 && onlyBr(p)); // paragraphs of line breaks only (3.2.5-C): not editable while the mode lasts, so nothing typed or composed can reach them (the browser would remove the placeholder break: a structure violation that threw the whole draft away)
  // Which original span holds which run: the DOM spans of a paragraph are its non-br runs in order (the committed text applied), key = the run's index in the FILE's paragraph, or 'fill' for the run made of text typed into an empty paragraph.
  let canCommit = !!(paras && orig && paras.length === paras0.length && orig.length === paras0.length);
  const slots = paras0.map((p, i) => {
    if (!canCommit) return [];
    const spans = [...p.children].filter((c) => c.classList.contains('sf-r') && !isPh(c)); const runs = []; // (the placeholder span is not a run: 3.2.5-C)
    paras[i].runs.forEach((r, ri) => { if (!r.br) runs.push(r.fill ? 'fill' : ri); });
    if (runs.length === spans.length) return spans.map((span, k) => ({ span, key: runs[k] }));
    canCommit = false; return [];
  });
  let live = true; let ended = false;

  const intact = () => {
    if ([...inner.childNodes].some((n) => n.nodeType !== 1 || !n.classList.contains('sf-p'))) return false;
    const ps = [...inner.children]; if (ps.length !== count0) return false;
    return ps.every((p, i) => {
      if (p.dir !== dirs0[i]) return false;
      const bu = bulletOf(p); const nb = p.querySelectorAll('br').length;
      if (nb !== brs0[i] && !(brs0[i] === 0 && nb === 1 && onlyBr(p))) return false; // (a line break is neither added nor removed; the one placeholder <br> of an emptied paragraph is not a line break)
      return (bu ? 1 : 0) === bullets0[i] && (!bu || p.firstChild === bu); // (a bullet is the very first node: nothing, not even a text node, before it)
    });
  };

  /** The range an input event is about: its target range, else the current selection. */
  function rangeOf(ev) {
    const t = ev.getTargetRanges && ev.getTargetRanges();
    if (t && t.length) { const r = doc.createRange(); try { r.setStart(t[0].startContainer, t[0].startOffset); r.setEnd(t[0].endContainer, t[0].endOffset); return r; } catch { /* fall through */ } }
    const s = doc.getSelection(); return s && s.rangeCount ? s.getRangeAt(0).cloneRange() : null;
  }
  /** Would an edit of this range change the structure: start or end outside one paragraph of this text, or start before the end of the paragraph's bullet. */
  function structural(r) {
    if (!r) return true;
    const a = paraOf(r.startContainer); const b = paraOf(r.endContainer);
    if (!a || a !== b || !inner.contains(a)) return true;
    if (brOnly[[...inner.children].indexOf(a)]) return true; // a paragraph of line breaks only
    const bu = bulletOf(a);
    if (bu) { const after = doc.createRange(); after.setStartAfter(bu); after.collapse(true); if (r.compareBoundaryPoints(Range.START_TO_START, after) < 0) return true; }
    if (!r.collapsed && brs0[[...inner.children].indexOf(a)] > 0) for (const br of a.querySelectorAll('br')) if (r.intersectsNode(br)) return true; // a selection never takes a line break with it
    return false;
  }
  /** Is there a line break right next to the caret in the direction of the key (dir < 0: Backspace, dir > 0: Delete)? Then the key would remove it. */
  function adjacentBr(r, dir) {
    const p = paraOf(r.startContainer); if (!p) return false;
    const c = dir < 0 ? r.startContainer : r.endContainer; const o = dir < 0 ? r.startOffset : r.endOffset; let n;
    if (c.nodeType === 3) { if (dir < 0 ? o > 0 : o < c.data.length) return false; n = c; } else { const k = c.childNodes[dir < 0 ? o - 1 : o]; if (k) return k.nodeName === 'BR'; n = c; }
    while (n && n !== p) { const sib = dir < 0 ? n.previousSibling : n.nextSibling; if (sib) return sib.nodeName === 'BR'; n = n.parentNode; }
    return false;
  }

  /** A collapsed caret with nothing to delete in the direction of the key inside its paragraph (dir < 0: Backspace, nothing after the bullet before it; dir > 0: Delete, nothing after it): the key would merge paragraphs or eat the bullet. */
  function atEdge(r, dir) {
    const p = paraOf(r.startContainer); const bu = bulletOf(p); const q = doc.createRange();
    if (dir < 0) { if (bu) q.setStartAfter(bu); else q.setStart(p, 0); q.setEnd(r.startContainer, r.startOffset); } else { q.setStart(r.endContainer, r.endOffset); q.setEnd(p, p.childNodes.length); }
    return q.toString() === '';
  }
  /**
   * Typing into a paragraph that has no text (emptied, or empty in the file) or over ALL of its text: the browser would carry the deleted text's style over with an inline `style` it writes into the attribute
   * text, which the page's CSP refuses (a console violation) and which would leave a formatting artefact in the draft. The editor puts the typed text into the paragraph's ORIGINAL first run instead, so it keeps
   * that run's formatting. Returns true when it did. (The browser's own undo of this one step is not available.)
   */
  function typeIntoWholeParagraph(data, r) {
    const p = paraOf(r.startContainer); const i = [...inner.children].indexOf(p); const src = firstRun[i]; if (!p || !src || brs0[i] > 0) return false; // (a paragraph with line breaks is never replaced as a whole)
    const bu = bulletOf(p); const q = doc.createRange(); q.selectNodeContents(p); if (bu) q.setStartAfter(bu); const all = q.toString();
    const empty = all === '' || (empty0[i] && all === ZWSP); const whole = !r.collapsed && r.toString() === all;
    if (!empty && !whole) return false;
    p.replaceChildren(...(bu ? [bu] : []), src); collapsed.add(p);
    let tn = [...src.childNodes].find((n) => n.nodeType === 3); if (!tn) { tn = doc.createTextNode(''); src.append(tn); }
    tn.data = data; const c = doc.createRange(); c.setStart(tn, tn.data.length); c.collapse(true); const s = doc.getSelection(); s.removeAllRanges(); s.addRange(c);
    return true;
  }
  const onBeforeInput = (ev) => {
    if (!ALLOWED.has(ev.inputType)) { ev.preventDefault(); return; } // paragraphs, line breaks, formatting, drop, links, ...
    if (ev.inputType === 'insertCompositionText' || ev.inputType === 'deleteCompositionText') return; // not cancelable: the structure check below catches anything it does
    const r = rangeOf(ev); const dir = /Backward$/.test(ev.inputType) ? -1 : /Forward$/.test(ev.inputType) ? 1 : 0;
    // some browsers give no target range for a collapsed delete: then the caret itself is tested against the edge of its paragraph
    if (structural(r) || (dir && r.collapsed && (adjacentBr(r, dir) || (!(ev.getTargetRanges && ev.getTargetRanges().length) && atEdge(r, dir))))) { ev.preventDefault(); return; }
    if (ev.inputType === 'insertText' && typeof ev.data === 'string' && ev.data && typeIntoWholeParagraph(ev.data, r)) ev.preventDefault();
  };
  /**
   * When ALL the text of a run is deleted the browser removes its span, and the next typing there makes it write a styled clone span (inline style text the page's CSP refuses; the new text would also have no run).
   * So a run that went missing while others of its paragraph are still there is put back, the original span, in its own place, holding one zero-width placeholder, and the caret goes into it:
   * what is typed lands in that run, with its formatting. (collect() leaves the placeholder out.) When nothing of the paragraph is left, typeIntoWholeParagraph() does the same job.
   */
  const placeholders = new Set(); const collapsed = new Set(); // collapsed = paragraphs typeIntoWholeParagraph() reduced to their first run on purpose: their other runs stay empty
  function reattach() {
    let caretTo = null;
    paras0.forEach((p, i) => {
      const sl = slots[i]; if (collapsed.has(p)) return; const alive = sl.map((s) => p.contains(s.span)); if (!sl.length || alive.every(Boolean) || !alive.some(Boolean)) return;
      sl.forEach((s, k) => {
        if (alive[k]) return;
        s.span.replaceChildren(doc.createTextNode(ZWSP)); placeholders.add(s.span);
        let ref = null; for (let j = k + 1; j < sl.length; j++) if (alive[j]) { ref = sl[j].span; break; }
        if (ref) ref.before(s.span); else { let prev = null; for (let j = k - 1; j >= 0; j--) if (alive[j]) { prev = sl[j].span; break; } prev.after(s.span); }
        alive[k] = true; if (!caretTo) caretTo = s.span.firstChild;
      });
    });
    if (caretTo) { const c = doc.createRange(); c.setStart(caretTo, caretTo.data.length); c.collapse(true); const s = doc.getSelection(); s.removeAllRanges(); s.addRange(c); }
  }
  const onInput = () => { reattach(); if (!intact()) end('structure'); };
  const onKeyDown = (ev) => {
    if (ev.key === 'Escape') return; // handled (captured) below
    if (ev.key === 'Enter' && !ev.isComposing) ev.preventDefault(); // no paragraph split, no line break
    ev.stopPropagation(); // the preview dialog turns the arrow keys into slide navigation
  };
  const onPaste = (ev) => {
    ev.preventDefault();
    const t = (ev.clipboardData && ev.clipboardData.getData('text/plain') || '').replace(/\r\n|[\r\n\u2028\u2029]/g, ' ');
    if (!t) return;
    const s = doc.getSelection(); const r = s && s.rangeCount ? s.getRangeAt(0) : null;
    if (structural(r)) return; // (execCommand does not give the input rules a cancelable event: the same rules are applied here)
    if (typeIntoWholeParagraph(t, r)) return;
    try { doc.execCommand('insertText', false, t); } catch { /* nothing to insert with */ }
  };
  const stop = (ev) => ev.preventDefault();
  const onEsc = (ev) => { if (ev.key === 'Escape') { ev.preventDefault(); ev.stopPropagation(); end('escape'); } }; // before the dialog sees it (it would close)
  const onFocusOut = () => { setTimeout(() => { if (live && doc.hasFocus() && !inner.contains(doc.activeElement)) end('blur'); }, 0); }; // focus left the text (not: the window lost focus)
  const onSelChange = () => { // a caret before a bullet (Home, arrow keys) is moved to just after it: nothing may be typed before a bullet
    const s = doc.getSelection(); if (!live || !s || !s.rangeCount || !s.isCollapsed || !inner.contains(s.anchorNode)) return;
    const r = s.getRangeAt(0); const p = paraOf(r.startContainer); const bu = p && bulletOf(p); if (!bu) return;
    const after = doc.createRange(); after.setStartAfter(bu); after.collapse(true);
    if (r.compareBoundaryPoints(Range.START_TO_START, after) < 0) { s.removeAllRanges(); s.addRange(after); }
  };
  const onBoxDown = (ev) => { if (ev.button === 0 && !inner.contains(ev.target)) ev.preventDefault(); }; // a press on the shape's padding keeps the caret

  for (const [a, v] of [['contenteditable', 'plaintext-only'], ['spellcheck', 'false'], ['autocapitalize', 'off'], ['autocorrect', 'off'], ['role', 'textbox'], ['aria-multiline', 'true']]) inner.setAttribute(a, v);
  inner.classList.add('sf-text-editing');
  for (const p of inner.children) { const bu = bulletOf(p); if (bu) bu.setAttribute('contenteditable', 'false'); }
  paras0.forEach((p, i) => { if (brOnly[i]) p.setAttribute('contenteditable', 'false'); }); // (put back by restore(), like the bullets')
  inner.addEventListener('beforeinput', onBeforeInput); inner.addEventListener('input', onInput); inner.addEventListener('keydown', onKeyDown);
  inner.addEventListener('paste', onPaste); inner.addEventListener('drop', stop); inner.addEventListener('dragover', stop); inner.addEventListener('dragstart', stop); inner.addEventListener('focusout', onFocusOut);
  box.addEventListener('mousedown', onBoxDown, true); doc.addEventListener('keydown', onEsc, true); doc.addEventListener('selectionchange', onSelChange);

  function placeCaret() {
    let r = null;
    if (point) {
      if (doc.caretPositionFromPoint) { const c = doc.caretPositionFromPoint(point.x, point.y); if (c) { r = doc.createRange(); r.setStart(c.offsetNode, c.offset); r.collapse(true); } }
      else if (doc.caretRangeFromPoint) r = doc.caretRangeFromPoint(point.x, point.y);
    }
    const p = r && paraOf(r.startContainer);
    if (!r || !p || !inner.contains(p) || (r.startContainer.nodeType === 1 ? r.startContainer : r.startContainer.parentElement).closest('.sf-bu') || structural(r)) { // not on editable text: the end of the last paragraph
      let li = count0 - 1; while (li > 0 && brOnly[li]) li--; /* the last paragraph that can be typed into */ const last = inner.children[li]; r = doc.createRange(); r.selectNodeContents(last); r.collapse(false);
    }
    const s = doc.getSelection(); s.removeAllRanges(); s.addRange(r);
  }
  // a double click has already selected the word under the pointer (the browser does that): keep it, so typing replaces it; otherwise the caret goes to the pointer
  const s0 = doc.getSelection(); const word = s0 && s0.rangeCount && !s0.isCollapsed && inner.contains(s0.anchorNode) && inner.contains(s0.focusNode) ? s0.getRangeAt(0).cloneRange() : null;
  inner.focus({ preventScroll: true });
  if (word && !structural(word)) { const s = doc.getSelection(); s.removeAllRanges(); s.addRange(word); } else placeCaret();

  /**
   * The text change this draft makes against the FILE's text: { paras: [ { p, runs: [ { r, text } ], fill? } ], skipped } (the changed runs only, see edit-state.commitText), or null when it cannot be read
   * safely (no paragraphs to map to, the structure is not intact). Each original span's text is its run's text; text that is not in an original span (a span the browser made, a bare text node) goes to the
   * original span before it, or, at the start, to the first one that is still there; an original span that is gone is an empty run (it stays a run). Text in an empty paragraph is `fill` (the placeholder
   * is not text). `skipped` = text that has no run to go to (a paragraph with only line breaks) was not kept. Nothing is normalised: NBSP, bidi marks and every other character are the draft's own.
   */
  function collect() {
    if (!canCommit || !intact()) return null;
    const out = []; let skipped = false;
    paras0.forEach((p, i) => {
      const sl = slots[i]; const tx = sl.map(() => ''); let lead = ''; let owner = -1;
      const walk = (nodes) => {
        for (const n of nodes) {
          if (n.nodeType === 3) { if (owner >= 0) tx[owner] += n.data; else lead += n.data; continue; }
          if (n.nodeType !== 1 || n.classList.contains('sf-bu') || n.tagName === 'BR') continue;
          if (isPh(n)) { const t = n.textContent.replace(ZWSP, ''); if (owner >= 0) tx[owner] += t; else lead += t; continue; } // (the placeholder: its zero-width character is not text; anything typed into it belongs to the run before it, else it is the paragraph's leading text)
          const si = sl.findIndex((s) => s.span === n);
          if (si >= 0) { owner = si; tx[si] += placeholders.has(n) ? n.textContent.replace(ZWSP, '') : n.textContent; continue; } // (a placeholder put back by reattach() is not text)
          walk(n.childNodes);
        }
      };
      walk(p.childNodes);
      const runs = []; let fill;
      if (sl.length) {
        if (lead) { const fs = sl.findIndex((s) => p.contains(s.span)); tx[fs >= 0 ? fs : 0] = lead + tx[fs >= 0 ? fs : 0]; }
        sl.forEach((s, k) => { if (s.key === 'fill') { if (tx[k] !== '') fill = tx[k]; } else if (tx[k] !== orig[i].runs[s.key].text) runs.push({ r: s.key, text: tx[k] }); });
      } else {
        const t = empty0[i] ? lead.replace(ZWSP, '') : lead;
        if (t !== '') { if (orig[i].runs.length === 0 && orig[i].defRun) fill = t; else skipped = true; }
      }
      if (runs.length || fill !== undefined) out.push({ p: i, runs, ...(fill !== undefined ? { fill } : {}) });
    });
    return { paras: out, skipped };
  }

  function end(reason = 'end') {
    if (!live) return; live = false;
    const draft = CANCEL.has(reason) ? null : collect(); // read while the draft is still in the DOM
    inner.removeEventListener('beforeinput', onBeforeInput); inner.removeEventListener('input', onInput); inner.removeEventListener('keydown', onKeyDown);
    inner.removeEventListener('paste', onPaste); inner.removeEventListener('drop', stop); inner.removeEventListener('dragover', stop); inner.removeEventListener('dragstart', stop); inner.removeEventListener('focusout', onFocusOut);
    box.removeEventListener('mousedown', onBoxDown, true); doc.removeEventListener('keydown', onEsc, true); doc.removeEventListener('selectionchange', onSelChange);
    const s = doc.getSelection(); if (s && s.rangeCount && inner.contains(s.anchorNode)) s.removeAllRanges();
    if (doc.activeElement === inner) inner.blur();
    restore(); // the original, as it was drawn
    for (const [a, v] of prior) { if (v === null) inner.removeAttribute(a); else inner.setAttribute(a, v); }
    inner.classList.remove('sf-text-editing');
    ended = true; onEnd?.(reason, draft);
  }

  return {
    inner,
    get active() { return live; },
    get ended() { return ended; },
    get canCommit() { return canCommit; },
    /** How many paragraphs of line breaks only were made non-editable (Phase 3.2.5-C). */
    get locked() { return brOnly.filter(Boolean).length; },
    collect,
    /** The draft, one string per paragraph, exactly as typed: bullets are not text, the placeholder of an empty paragraph and the browser's placeholder <br> are left out, nothing else is changed. */
    text() {
      return [...inner.children].map((p, i) => {
        let s = ''; const kids = [...p.childNodes].filter((n) => !(n.nodeType === 1 && n.classList.contains('sf-bu')));
        const walk = (nodes) => { for (const c of nodes) { if (c.nodeType === 3) s += c.data; else if (c.nodeType === 1) { if (c.tagName === 'BR') s += kids.length === 1 ? '' : '\n'; else if (placeholders.has(c)) s += c.textContent.replace(ZWSP, ''); else walk(c.childNodes); } } };
        walk(kids);
        return empty0[i] ? s.replace(ZWSP, '') : s;
      });
    },
    end,
  };
}
