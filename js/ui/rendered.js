// "Rendered" tab (Phase 3.1): a read-only, native rendering of the selected slide (real text, vector shapes, tables, pictures), plus an honest
// support report. It only reads the parsed deck: nothing is edited, saved or written to the template, and the other tabs are unaffected.
// The slide shown is always the one selected in the Slides tab: preview.js calls goto(i) from its own select().
import { h } from '../lib/dom.js';
import { plural } from '../lib/format.js';
import { buildScene } from '../model/scene.js';
import { renderSlide } from '../render/render.js';
import { attachSelection } from '../render/selection.js';
import { editStateOf } from '../model/edit-state.js';
import { isBlankText } from '../model/textbox.js';
import { loadStyles } from '../render/fonts.js';
import { styleBar } from './style-bar.js';
import { errorState, friendly, spinner } from './states.js';

const LEVEL = { native: 'Native', approximated: 'Approximated', opaque: 'Not drawn' };
const KIND = { shape: 'Shape', image: 'Picture', line: 'Line', group: 'Group', table: 'Table', chart: 'Chart', diagram: 'SmartArt', ole: 'Embedded object', frame: 'Content', content: 'Content', unknown: 'Unsupported', alt: 'Alternate content' };

/** Returns { el, goto(i), show(), hide(), dispose() }. `onPick(i)` asks the host to change the selected slide. */
export function renderedPanel(parsed, { onPick, current = () => 0 } = {}) {
  const scenes = new Map();
  const edits = editStateOf(parsed); // Phase 3.2.3-B / 3.2.4-B: completed moves and completed resizes, kept in memory for as long as this deck is open (temporary: nothing is written to the file or saved)
  let handle = null; let token = 0; let visible = false; let index = -1; let ro = null;
  let autoEdit = null; let pendingNote = null; // Phase 3.2.7-A: the id of a text box just made (text edit mode starts when the slide has been drawn with it), and a status line to show once the slide is drawn again
  let sel = null; let picked = null; // Phase 3.2.1: the selection controller of the drawn slide, and { slide, id } of the selected element (kept when the same slide is drawn again, dropped on any other slide)

  const film = h('ol', { class: 'film', 'aria-label': 'Slides (rendered)' }, parsed.slides.map((s, i) => h('li', {}, h('button', { type: 'button', class: 'film-item', 'data-i': i, onclick: () => onPick?.(i) },
    h('span', { class: 'film-n' }, i + 1), h('span', { class: 'film-t' }, s.title || 'No title'), s.hidden && h('span', { class: 'tag' }, 'Hidden')))));
  const frame = h('div', { class: 'sf-frame' });
  const note = h('p', { class: 'muted fine' });
  const selNote = h('p', { class: 'muted fine sf-sel-status', role: 'status', 'aria-live': 'polite' });
  const details = h('div', { class: 'details' });
  // Phase 3.2.7-A/B: what the preview can add to a slide (a text box and five shapes) and the button that removes an added element
  const ADD = [['textbox', 'Add text box'], ['rect', 'Rectangle'], ['roundRect', 'Rounded rectangle'], ['ellipse', 'Ellipse'], ['line', 'Line'], ['arrow', 'Arrow']];
  const addBtns = ADD.map(([kind, label]) => h('button', { type: 'button', class: `btn ${kind === 'textbox' ? 'sf-add-tb' : 'sf-add-sh'}`, 'data-act': kind === 'textbox' ? 'add-textbox' : 'add-shape', 'data-shape': kind, disabled: true, onclick: () => addElement(kind) }, label));
  const removeBtn = h('button', { type: 'button', class: 'btn sf-add-rm', 'data-act': 'remove-added', disabled: true, onclick: () => removeSelected() }, 'Remove');
  const addOff = (off) => { for (const b of addBtns) b.disabled = off; if (off) removeBtn.disabled = true; };
  const bar = styleBar({ edits }); // Phase 3.2.6-D: the style controls of the selected element (they call the controller's style(); the slide's notes are refreshed after every change)
  const el = h('div', { class: 'preview rendered', 'data-rendered': 'true' },
    film,
    h('div', { class: 'stage-wrap' }, frame, h('div', { class: 'sf-add' }, ...addBtns), h('div', { class: 'sf-add sf-add-sub' }, removeBtn, h('span', { class: 'sf-add-hint muted' }, 'Adds a text box or a shape to this slide; Remove deletes an element you added. Preview only: nothing is saved.')), bar.el, note, selNote),
    details);

  const mark = (i) => [...film.children].forEach((li, k) => { const b = li.firstChild; b.setAttribute('aria-current', String(k === i)); if (k === i && visible) b.scrollIntoView?.({ block: 'nearest', inline: 'nearest' }); });

  function describe(rep, fonts, scene) {
    const total = rep.total || 0;
    const badge = (lv) => h('span', { class: 'sf-badge', 'data-support': lv }, LEVEL[lv]);
    const hidden = scene.all.filter((e) => e.hidden).length;
    const issues = rep.issues;
    return [
      h('h3', {}, `Slide ${index + 1}, rendered`),
      h('p', { class: 'muted' }, 'Drawn from the file as live text and vector shapes. Read-only: nothing here is saved.'),
      h('h4', {}, 'Support on this slide'),
      h('ul', { class: 'file-list' },
        ['native', 'approximated', 'opaque'].map((lv) => h('li', {}, badge(lv), ` ${plural(rep.counts[lv], 'element')}`))),
      total === 0 ? h('p', { class: 'muted' }, 'This slide has no elements.') : null,
      rep.background !== 'native' ? h('p', { class: 'muted' }, 'The background fill is approximated.') : null,
      hidden ? h('p', { class: 'muted' }, `${plural(hidden, 'hidden element')} in the file ${hidden === 1 ? 'is' : 'are'} kept but not shown, as in PowerPoint.`) : null,
      rep.unplaced.length ? h('p', { class: 'muted' }, `${plural(rep.unplaced.length, 'element')} ${rep.unplaced.length === 1 ? 'has' : 'have'} no position that can be resolved and ${rep.unplaced.length === 1 ? 'is' : 'are'} not drawn. Kept in the file and in the model.`) : null,
      issues.length ? h('details', { class: 'notice' }, h('summary', {}, `${plural(issues.length, 'element')} not drawn exactly`),
        h('ul', { class: 'file-list' }, issues.map((i) => h('li', {}, badge(i.support), ' ', `${KIND[i.kind] || i.kind}${i.name ? ` “${i.name}”` : ''}`, h('ul', {}, i.reasons.map((r) => h('li', {}, r))))))) : null,
      fonts.length ? h('h4', {}, 'Fonts') : null,
      fonts.length ? h('ul', { class: 'file-list' }, fonts.map((f) => h('li', {}, f.name, ': ', f.original ? 'available on this device' : f.substitute ? `shown with the equivalent ${f.substitute.replace(/^SF /, '')}` : 'not installed here, shown with a generic font', f.arabic ? ' (Arabic text)' : ''))) : null,
    ].filter(Boolean);
  }

  const kindOf = (e) => (e && e.added === 'textbox' ? 'Text box' : e ? KIND[e.kind] || e.kind : 'the element'); // Phase 3.2.7-A
  const IDLE = 'Nothing selected. Click or tap an element to select it, then drag it to move it (preview only: nothing is edited or saved).';
  const pt = (n) => `${Math.round(n * 10) / 10}`;
  /** Status line. `mv` = the selected element's temporary preview offset in slide points (Phase 3.2.3-A), shown only while it is not zero; `rs` = its temporary resize preview (Phase 3.2.4-A), shown only while its size differs from the file's. */
  function say(scene, id, mv, rs) {
    const e = id && scene?.byId.get(id);
    const moved = e && mv && (mv.dx || mv.dy) ? `. Moved in the preview by ${pt(mv.dx)} pt, ${pt(mv.dy)} pt (not saved)` : '';
    const sized = e && rs && (rs.dw || rs.dh) ? `. Resized in the preview to ${pt(rs.w)} × ${pt(rs.h)} pt (not saved)` : '';
    selNote.textContent = e ? `Selected: ${kindOf(e)}${e.name ? ` “${e.name}”` : ''} (${e.id})${moved}${sized}` : scene ? IDLE : '';
  }

  async function draw(i) {
    const my = ++token;
    mark(i); addOff(true);
    sel?.dispose(); sel = null; bar.bind(null); selNote.textContent = ''; // the old slide's selection ends with the old slide
    if (picked && picked.slide !== i) picked = null;
    frame.replaceChildren(h('p', { class: 'status' }, spinner(), 'Rendering slide…'));
    try {
      await loadStyles();
      const part = parsed.slides[i].part; const sg = JSON.stringify(edits.addedFor(part)); // Phase 3.2.7-A: a scene built before the slide's text boxes changed is built again
      let scene = scenes.get(i); if (scene && scene.addedSig !== sg) scene = null;
      if (!scene) { scene = await buildScene(parsed, i, { added: edits.addedFor(part) }); scene.addedSig = sg; scenes.set(i, scene); }
      if (my !== token) return;
      handle?.dispose(); ro?.disconnect();
      const holder = h('div', { class: 'sf-holder' });
      frame.replaceChildren(holder);
      handle = await renderSlide(scene, holder);
      if (my !== token) { handle.dispose(); return; }
      handle.root.setAttribute('role', 'group'); handle.root.setAttribute('aria-label', `Slide ${i + 1}, rendered from the file (read-only)`);
      // The count line and the details of the slide, from the scene as it is NOW: text fitting adds notes (and, since 3.2.5-C, takes them off again), so this runs after the first drawing and after every committed text edit.
      const refresh = () => {
        handle.fitText();
        const rep = handle.report();
        const made = scene.elements.filter((e) => e.added).length;
        note.textContent = `${plural(rep.counts.native + rep.counts.approximated + rep.counts.opaque, 'element')} drawn from the file${made ? ` (${plural(made, 'new element')} included)` : ''}: ${rep.counts.native} native, ${rep.counts.approximated} approximated, ${rep.counts.opaque} not drawn.`;
        details.replaceChildren(...describe(rep, rep.fonts, scene));
      };
      let pendingFit = false; // a resize of the frame while text is being edited is applied when the mode ends (re-fitting text under the caret would make it jump)
      const fit = () => { if (sel && sel.editing != null) { pendingFit = true; return; } handle.setWidth(frame.clientWidth); };
      fit();
      if (typeof ResizeObserver !== 'undefined') { ro = new ResizeObserver(() => fit()); ro.observe(frame); }
      sel = attachSelection(handle, scene, {
        onChange: (id) => { const prev = picked && picked.slide === i ? picked.id : null; if (prev && prev !== id) emptyLeft(i, prev); picked = id ? { slide: i, id } : null; removeBtn.disabled = !(id && edits.isAdded(id)); say(scene, id, id && sel ? sel.offset(id) : null, id && sel ? sel.size(id) : null); bar.sync(); },
        onMove: (m) => { if (m.phase !== 'start') say(scene, m.id, m, sel ? sel.size(m.id) : null); }, // a completed move is committed to the temporary edit-state, which draws it again on every redraw of this slide
        onResize: (r) => { if (r.phase !== 'start') say(scene, r.id, sel ? sel.offset(r.id) : null, r); }, // a completed resize is committed to the temporary edit-state (3.2.4-B), which draws it again on every redraw of this slide
        onEdit: (m) => { // Phase 3.2.5-A / B: text edit mode (a draft kept in the temporary edit-state on every exit but Escape; nothing is saved)
          const now = () => say(scene, sel ? sel.id : null, sel && sel.id ? sel.offset(sel.id) : null, sel && sel.id ? sel.size(sel.id) : null);
          if (m.phase === 'start') selNote.textContent = `Editing the text of ${kindOf(scene.byId.get(m.id))}${scene.byId.get(m.id)?.name ? ` “${scene.byId.get(m.id).name}”` : ''} in the preview. Esc cancels; a click outside keeps the change in memory (nothing is saved).${m.locked ? ` ${m.locked === 1 ? 'A paragraph that holds only a line break cannot' : 'Paragraphs that hold only line breaks cannot'} be typed into yet.` : ''}`;
          else if (m.phase === 'refused') { now(); selNote.textContent += `. ${m.message}`; }
          else { now(); if (m.added) afterBox(i, m); if (m.committed && m.changed) selNote.textContent += '. Text changed in the preview (kept while this deck is open; nothing is saved).'; if (m.skipped) selNote.textContent += ' Some typed text had no place to go and was not kept.'; if (pendingFit) { pendingFit = false; fit(); } if (m.committed && m.changed) refresh(); }
        },
        edits,
      });
      bar.bind({ sel, scene, onApplied: () => { if (my === token) refresh(); } });
      if (!(picked && sel.select(picked.id))) { picked = null; say(scene, null); } // drawn again (tab shown again): the same element stays selected
      refresh(); addOff(false); removeBtn.disabled = !(sel.id && edits.isAdded(sel.id));
      if (pendingNote) { selNote.textContent = pendingNote; pendingNote = null; }
      if (autoEdit && sel.id === autoEdit) { const id = autoEdit; autoEdit = null; startEdit(my, sel, id, 0); } // a text box just made: its text is edited at once
    } catch (e) {
      console.error(e);
      if (my !== token) return;
      frame.replaceChildren(errorState(new Error(`This slide could not be rendered. ${friendly(e)}`), () => draw(i)));
      note.textContent = ''; selNote.textContent = '';
      details.replaceChildren();
    }
  }

  /** Phase 3.2.7-A/B: make a new text box (its text is edited at once) or shape ('rect', 'roundRect', 'ellipse', 'line', 'arrow') on the drawn slide and select it. */
  function addElement(kind) {
    const i = index; const scene = scenes.get(i);
    if (!visible || !handle || !scene || addBtns[0].disabled) return;
    sel?.endEdit('api'); // a box being edited keeps its draft
    const r = edits.addElement(scene, kind); if (!r.ok) return;
    picked = { slide: i, id: r.id }; autoEdit = kind === 'textbox' ? r.id : null;
    draw(i); // the scene is built again with the new box in it
  }
  // A press in the panel (the slide, the style bar, the button) is finished before a redraw caused by a new box's edit begins: that redraw would disable and unbind the controls under the pointer.
  let pressing = false; const afterPress = [];
  const onDown = () => { pressing = true; };
  const onUp = () => { setTimeout(() => { pressing = false; for (const f of afterPress.splice(0)) f(); }, 0); };
  el.addEventListener('pointerdown', onDown, true); document.addEventListener('pointerup', onUp, true); document.addEventListener('pointercancel', onUp, true);
  /** Remove the selected element when it is one the preview added (the file's own elements are never removed). */
  function removeSelected() {
    const id = sel && sel.id; if (!id || !edits.isAdded(id)) return;
    const name = scenes.get(index)?.byId.get(id)?.name; sel.endEdit('api');
    if (!edits.removeAdded(id)) return;
    picked = null; autoEdit = null; pendingNote = `Removed ${name ? `“${name}”` : 'the element'}.`;
    draw(index);
  }
  /** Text edit mode of a box just made: the styles of the slide are still being drawn for a moment after it is attached (the controller refuses edit mode meanwhile), so it is tried again shortly. */
  function startEdit(my, ctl, id, n) {
    if (my !== token || sel !== ctl || ctl.id !== id) return;
    const r = ctl.beginEdit(id);
    if (!r.ok && r.reason === 'style-pending' && n < 40) setTimeout(() => startEdit(my, ctl, id, n + 1), 50);
  }
  /** The selection left a new text box (slide i): one that never got any text is removed. */
  function emptyLeft(i, id) { const r = edits.getAdded(id); if (r && r.paras && isBlankText(r.paras.map((p) => p.text))) afterBox(i, { id, empty: true }); }
  /** A new text box's text edit ended (slide i): an empty box is removed, a box whose text changed is drawn again from its record. Runs after the press that ended the edit is over (that press may select another element). */
  function afterBox(i, m) {
    const run = () => {
      if (m.empty && !edits.removeAdded(m.id)) return; // (already removed by an earlier call: nothing more to do)
      if (m.empty) { if (picked && picked.id === m.id) picked = null; if (autoEdit === m.id) autoEdit = null; if (visible && index === i) pendingNote = 'Empty text box removed.'; }
      if (visible && index === i && scenes.get(i)?.addedSig !== JSON.stringify(edits.addedFor(parsed.slides[i].part))) draw(i);
      else if (pendingNote && visible && index === i) { selNote.textContent = pendingNote; pendingNote = null; }
    };
    setTimeout(() => (pressing ? afterPress.push(run) : run()), 0);
  }

  return {
    el,
    /** The host selected slide i. Draws it when the tab is visible, otherwise on the next show(). */
    goto(i) { index = i; mark(i); if (visible) draw(i); },
    show() { visible = true; el.hidden = false; if (index < 0) index = current(); mark(index); draw(index); },
    hide() { sel?.endEdit('hide'); visible = false; token++; },
    dispose() { token++; afterPress.length = 0; el.removeEventListener('pointerdown', onDown, true); document.removeEventListener('pointerup', onUp, true); document.removeEventListener('pointercancel', onUp, true); ro?.disconnect(); bar.dispose(); sel?.dispose(); sel = null; handle?.dispose(); handle = null; scenes.clear(); edits.clear(); },
  };
}
