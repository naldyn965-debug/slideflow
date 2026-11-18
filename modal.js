import { h, icon } from '../lib/dom.js';

/** Native <dialog>: focus trap and Esc come for free. Set m.busy = true to block closing during uploads. */
export function modal({ title, wide = false }) {
  const heading = h('h2', {}, title);
  const body = h('div', { class: 'modal-body' });
  const m = { busy: false, body, onClose: null, setTitle: (t) => { heading.textContent = t; }, close: () => dlg.close() };
  const dlg = h('dialog', { class: wide ? 'modal wide' : 'modal', 'aria-label': title },
    h('header', { class: 'modal-head' }, heading,
      h('button', { class: 'icon-btn', type: 'button', 'aria-label': 'Close', onclick: () => { if (!m.busy) dlg.close(); } }, icon('x'))),
    body);
  dlg.addEventListener('cancel', (e) => { if (m.busy) e.preventDefault(); });
  dlg.addEventListener('mousedown', (e) => { if (e.target === dlg && !m.busy) dlg.close(); });
  dlg.addEventListener('close', () => { m.onClose?.(); dlg.remove(); });
  document.body.append(dlg);
  dlg.showModal();
  m.dlg = dlg;
  return m;
}

export function confirmDialog({ title, body, confirmLabel, danger = false }) {
  return new Promise((resolve) => {
    const m = modal({ title });
    let answer = false;
    m.onClose = () => resolve(answer);
    m.body.append(
      h('p', { class: 'muted' }, body),
      h('div', { class: 'actions' },
        h('button', { class: 'btn', type: 'button', onclick: () => m.close() }, 'Cancel'),
        h('button', { class: `btn ${danger ? 'btn-danger' : 'btn-primary'}`, type: 'button', onclick: () => { answer = true; m.close(); } }, confirmLabel)));
  });
}
