import { h, icon, clear } from '../lib/dom.js';

export function friendly(err) {
  const c = err?.code || '';
  if (c.includes('permission-denied') || c === 'storage/unauthorized') return "Access was denied. Check that you're signed in and that the Firestore and Storage rules are deployed.";
  if (c.includes('unavailable') || c === 'storage/retry-limit-exceeded') return 'Could not reach the server. Check your connection and try again.';
  if (c === 'storage/quota-exceeded') return 'The Firebase Storage quota for this project is used up.';
  return err?.message || 'Unknown error.';
}

export const spinner = () => h('span', { class: 'spinner', 'aria-hidden': 'true' });

export const skeleton = (n = 3) => h('div', { class: 'skeleton-list', 'aria-busy': 'true', 'aria-label': 'Loading' }, Array.from({ length: n }, () => h('div', { class: 'skeleton' })));

export function emptyState({ title, body, actions = [] }) {
  return h('div', { class: 'empty' }, h('h3', {}, title), body && h('p', { class: 'muted' }, body), actions.length > 0 && h('div', { class: 'actions-left' }, ...actions));
}

export function errorState(err, retry) {
  return h('div', { class: 'error-panel', role: 'alert' }, icon('alert'),
    h('div', {}, h('strong', {}, 'Something went wrong'), h('p', { class: 'muted' }, friendly(err)),
      retry && h('button', { class: 'btn', type: 'button', onclick: retry }, 'Try again')));
}

/** Skeleton -> data -> render, with an error panel and retry. Returns the reload function. */
export function loadInto(root, load, render) {
  const run = async () => {
    clear(root).append(skeleton());
    try { const data = await load(); clear(root); render(data); }
    catch (e) { console.error(e); clear(root).append(errorState(e, run)); }
  };
  run();
  return run;
}

export function pageHeader({ title, sub, actions = [] }) {
  return h('header', { class: 'page-head' }, h('div', {}, h('h1', {}, title), sub && h('p', { class: 'muted' }, sub)), h('div', { class: 'actions-left' }, ...actions));
}
