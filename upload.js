// Upload flow: pick -> inspect (validate + parse) -> review -> save. Nothing is stored until the user confirms.
import { h } from '../lib/dom.js';
import { LIMITS } from '../config.js';
import { fmtBytes, plural, sizeLabel, stripExt } from '../lib/format.js';
import { inspectPptx } from '../pptx/validate.js';
import { summarize } from '../pptx/parser.js';
import { modal } from './modal.js';
import { errorState, spinner } from './states.js';

const ACCEPT = '.pptx,application/vnd.openxmlformats-officedocument.presentationml.presentation';

export function pickAndUpload(ctx, onSaved) {
  const input = h('input', { type: 'file', accept: ACCEPT });
  input.addEventListener('change', () => input.files[0] && openUpload(ctx, input.files[0], onSaved));
  input.click();
}

export async function openUpload(ctx, file, onSaved) {
  const m = modal({ title: 'Add template' });
  const show = (...nodes) => m.body.replaceChildren(...nodes);
  const status = (text) => h('p', { class: 'status' }, spinner(), text);

  m.busy = true;
  show(status(`Checking ${file.name}…`));
  let parsed;
  try { parsed = await inspectPptx(file); }
  catch (e) {
    m.busy = false;
    return show(errorState(e), h('div', { class: 'actions' },
      h('button', { class: 'btn', type: 'button', onclick: () => m.close() }, 'Close'),
      h('button', { class: 'btn btn-primary', type: 'button', onclick: () => { m.close(); pickAndUpload(ctx, onSaved); } }, 'Choose another file')));
  }
  m.busy = false;
  m.onClose = () => { parsed = null; };

  const review = (error) => {
    const name = h('input', { class: 'input', id: 'tpl-name', maxlength: LIMITS.maxNameLength, value: stripExt(file.name).slice(0, LIMITS.maxNameLength), required: true });
    const bar = h('progress', { max: 1, value: 0, class: 'hidden' });
    const save = h('button', { class: 'btn btn-primary', type: 'submit' }, 'Save to library');
    const cancel = h('button', { class: 'btn', type: 'button', onclick: () => m.close() }, 'Cancel');
    const facts = [['Slides', String(parsed.slideCount)], ['Slide size', sizeLabel(parsed.width, parsed.height)], ['Pictures', `${plural(parsed.imageRefs, 'reference')} to ${plural(parsed.mediaCount, 'media file')}`], ['File size', fmtBytes(file.size)]];

    const form = h('form', { class: 'stack', novalidate: true, onsubmit: async (e) => {
      e.preventDefault();
      const clean = name.value.trim();
      if (!clean) { name.focus(); return; }
      m.busy = true; save.disabled = cancel.disabled = name.disabled = true; bar.classList.remove('hidden');
      try {
        const t = await ctx.store.templates.add(file, { name: clean, summary: summarize(parsed) }, (p) => { bar.value = p; });
        m.busy = false; m.close(); ctx.toast(`Saved “${clean}” to your library.`, 'ok'); onSaved?.(t);
      } catch (err) { console.error(err); m.busy = false; review(err); }
    } },
      error && errorState(error),
      h('label', { class: 'field' }, h('span', {}, 'Template name'), name),
      h('dl', { class: 'facts' }, facts.flatMap(([k, v]) => [h('dt', {}, k), h('dd', {}, v)])),
      parsed.warnings.length > 0 && h('div', { class: 'notice' }, h('strong', {}, 'Heads up. You can still save this template.'), h('ul', {}, parsed.warnings.slice(0, 5).map((w) => h('li', {}, w)), parsed.warnings.length > 5 && h('li', {}, `and ${parsed.warnings.length - 5} more`))),
      bar,
      h('div', { class: 'actions' }, cancel, save));
    show(form);
    name.select();
  };
  review();
}
