// "Rendered" tab (Phase 3.1): a read-only, native rendering of the selected slide (real text, vector shapes, tables, pictures), plus an honest
// support report. It only reads the parsed deck: nothing is edited, saved or written to the template, and the other tabs are unaffected.
// The slide shown is always the one selected in the Slides tab: preview.js calls goto(i) from its own select().
import { h } from '../lib/dom.js';
import { plural } from '../lib/format.js';
import { buildScene } from '../model/scene.js';
import { renderSlide } from '../render/render.js';
import { attachSelection } from '../render/selection.js';
import { editStateOf } from '../model/edit-state.js';
import { loadStyles } from '../render/fonts.js';
import { errorState, friendly, spinner } from './states.js';

const LEVEL = { native: 'Native', approximated: 'Approximated', opaque: 'Not drawn' };
const KIND = { shape: 'Shape', image: 'Picture', line: 'Line', group: 'Group', table: 'Table', chart: 'Chart', diagram: 'SmartArt', ole: 'Embedded object', frame: 'Content', content: 'Content', unknown: 'Unsupported', alt: 'Alternate content' };

/** Returns { el, goto(i), show(), hide(), dispose() }. `onPick(i)` asks the host to change the selected slide. */
export function renderedPanel(parsed, { onPick, current = () => 0 } = {}) {
  const scenes = new Map();
  const edits = editStateOf(parsed); // Phase 3.2.3-B / 3.2.4-B: completed moves and completed resizes, kept in memory for as long as this deck is open (temporary: nothing is written to the file or saved)
  let handle = null; let token = 0; let visible = false; let index = -1; let ro = null;
  let sel = null; let picked = null; // Phase 3.2.1: the selection controller of the drawn slide, and { slide, id } of the selected element (kept when the same slide is drawn again, dropped on any other slide)

  const film = h('ol', { class: 'film', 'aria-label': 'Slides (rendered)' }, parsed.slides.map((s, i) => h('li', {}, h('button', { type: 'button', class: 'film-item', 'data-i': i, onclick: () => onPick?.(i) },
    h('span', { class: 'film-n' }, i + 1), h('span', { class: 'film-t' }, s.title || 'No title'), s.hidden && h('span', { class: 'tag' }, 'Hidden')))));
  const frame = h('div', { class: 'sf-frame' });
  const note = h('p', { class: 'muted fine' });
  const selNote = h('p', { class: 'muted fine sf-sel-status', role: 'status', 'aria-live': 'polite' });
  const details = h('div', { class: 'details' });
  const el = h('div', { class: 'preview rendered', 'data-rendered': 'true' },
    film,
    h('div', { class: 'stage-wrap' }, frame, note, selNote),
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

  const IDLE = 'Nothing selected. Click or tap an element to select it, then drag it to move it (preview only: nothing is edited or saved).';
  const pt = (n) => `${Math.round(n * 10) / 10}`;
  /** Status line. `mv` = the selected element's temporary preview offset in slide points (Phase 3.2.3-A), shown only while it is not zero; `rs` = its temporary resize preview (Phase 3.2.4-A), shown only while its size differs from the file's. */
  function say(scene, id, mv, rs) {
    const e = id && scene?.byId.get(id);
    const moved = e && mv && (mv.dx || mv.dy) ? `. Moved in the preview by ${pt(mv.dx)} pt, ${pt(mv.dy)} pt (not saved)` : '';
    const sized = e && rs && (rs.dw || rs.dh) ? `. Resized in the preview to ${pt(rs.w)} × ${pt(rs.h)} pt (not saved)` : '';
    selNote.textContent = e ? `Selected: ${KIND[e.kind] || e.kind}${e.name ? ` “${e.name}”` : ''} (${e.id})${moved}${sized}` : scene ? IDLE : '';
  }

  async function draw(i) {
    const my = ++token;
    mark(i);
    sel?.dispose(); sel = null; selNote.textContent = ''; // the old slide's selection ends with the old slide
    if (picked && picked.slide !== i) picked = null;
    frame.replaceChildren(h('p', { class: 'status' }, spinner(), 'Rendering slide…'));
    try {
      await loadStyles();
      let scene = scenes.get(i);
      if (!scene) { scene = await buildScene(parsed, i); scenes.set(i, scene); }
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
        note.textContent = `${plural(rep.counts.native + rep.counts.approximated + rep.counts.opaque, 'element')} drawn from the file: ${rep.counts.native} native, ${rep.counts.approximated} approximated, ${rep.counts.opaque} not drawn.`;
        details.replaceChildren(...describe(rep, rep.fonts, scene));
      };
      let pendingFit = false; // a resize of the frame while text is being edited is applied when the mode ends (re-fitting text under the caret would make it jump)
      const fit = () => { if (sel && sel.editing != null) { pendingFit = true; return; } handle.setWidth(frame.clientWidth); };
      fit();
      if (typeof ResizeObserver !== 'undefined') { ro = new ResizeObserver(() => fit()); ro.observe(frame); }
      sel = attachSelection(handle, scene, {
        onChange: (id) => { picked = id ? { slide: i, id } : null; say(scene, id, id && sel ? sel.offset(id) : null, id && sel ? sel.size(id) : null); },
        onMove: (m) => { if (m.phase !== 'start') say(scene, m.id, m, sel ? sel.size(m.id) : null); }, // a completed move is committed to the temporary edit-state, which draws it again on every redraw of this slide
        onResize: (r) => { if (r.phase !== 'start') say(scene, r.id, sel ? sel.offset(r.id) : null, r); }, // a completed resize is committed to the temporary edit-state (3.2.4-B), which draws it again on every redraw of this slide
        onEdit: (m) => { // Phase 3.2.5-A / B: text edit mode (a draft kept in the temporary edit-state on every exit but Escape; nothing is saved)
          const now = () => say(scene, sel ? sel.id : null, sel && sel.id ? sel.offset(sel.id) : null, sel && sel.id ? sel.size(sel.id) : null);
          if (m.phase === 'start') selNote.textContent = `Editing the text of ${KIND[scene.byId.get(m.id)?.kind] || 'the element'}${scene.byId.get(m.id)?.name ? ` “${scene.byId.get(m.id).name}”` : ''} in the preview. Esc cancels; a click outside keeps the change in memory (nothing is saved).${m.locked ? ` ${m.locked === 1 ? 'A paragraph that holds only a line break cannot' : 'Paragraphs that hold only line breaks cannot'} be typed into yet.` : ''}`;
          else if (m.phase === 'refused') { now(); selNote.textContent += `. ${m.message}`; }
          else { now(); if (m.committed && m.changed) selNote.textContent += '. Text changed in the preview (kept while this deck is open; nothing is saved).'; if (m.skipped) selNote.textContent += ' Some typed text had no place to go and was not kept.'; if (pendingFit) { pendingFit = false; fit(); } if (m.committed && m.changed) refresh(); }
        },
        edits,
      });
      if (!(picked && sel.select(picked.id))) { picked = null; say(scene, null); } // drawn again (tab shown again): the same element stays selected
      refresh();
    } catch (e) {
      console.error(e);
      if (my !== token) return;
      frame.replaceChildren(errorState(new Error(`This slide could not be rendered. ${friendly(e)}`), () => draw(i)));
      note.textContent = ''; selNote.textContent = '';
      details.replaceChildren();
    }
  }

  return {
    el,
    /** The host selected slide i. Draws it when the tab is visible, otherwise on the next show(). */
    goto(i) { index = i; mark(i); if (visible) draw(i); },
    show() { visible = true; el.hidden = false; if (index < 0) index = current(); mark(index); draw(index); },
    hide() { sel?.endEdit('hide'); visible = false; token++; },
    dispose() { token++; ro?.disconnect(); sel?.dispose(); sel = null; handle?.dispose(); handle = null; scenes.clear(); edits.clear(); },
  };
}
