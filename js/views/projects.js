import { h, icon } from '../lib/dom.js';
import { LIMITS } from '../config.js';
import { ago, MODE_LABEL } from '../lib/format.js';
import { confirmDialog, modal } from '../ui/modal.js';
import { emptyState, errorState, loadInto, pageHeader, spinner } from '../ui/states.js';

export async function openNewProject(ctx) {
  const m = modal({ title: 'New project' });
  m.body.append(h('p', { class: 'status' }, spinner(), 'Loading templates…'));
  let templates;
  try { templates = await ctx.store.templates.list(); }
  catch (e) { return m.body.replaceChildren(errorState(e)); }

  const name = h('input', { class: 'input', id: 'p-name', maxlength: LIMITS.maxNameLength, placeholder: 'Quarterly review', required: true });
  const tpl = h('select', { class: 'input' }, h('option', { value: '' }, 'No template'), templates.map((t) => h('option', { value: t.id }, t.name)));
  const mode = h('select', { class: 'input' }, Object.entries(MODE_LABEL).map(([v, l]) => h('option', { value: v }, l)));
  tpl.addEventListener('change', () => { mode.value = templates.find((t) => t.id === tpl.value)?.mode || 'preserve'; }); // default to the template's chosen mode
  const save = h('button', { class: 'btn btn-primary', type: 'submit' }, 'Create project');
  const err = h('div');
  m.body.replaceChildren(h('form', { class: 'stack', novalidate: true, onsubmit: async (e) => {
    e.preventDefault();
    const clean = name.value.trim();
    if (!clean) return name.focus();
    save.disabled = true; m.busy = true;
    try {
      const p = await ctx.store.projects.create({ name: clean, template: templates.find((t) => t.id === tpl.value), mode: mode.value });
      m.busy = false; m.close(); location.hash = `#/projects/${p.id}`;
    } catch (ex) { console.error(ex); m.busy = false; save.disabled = false; err.replaceChildren(errorState(ex)); }
  } },
    h('label', { class: 'field' }, h('span', {}, 'Project name'), name),
    h('label', { class: 'field' }, h('span', {}, 'Template'), tpl, templates.length === 0 && h('small', { class: 'muted' }, 'Upload a template first if you want to start from one.')),
    h('label', { class: 'field' }, h('span', {}, 'Design mode'), mode, h('small', { class: 'muted' }, 'Preserve Design keeps the template as is. Smart Enhance keeps its identity but improves readability. Used when slides are generated.')),
    err,
    h('div', { class: 'actions' }, h('button', { class: 'btn', type: 'button', onclick: () => m.close() }, 'Cancel'), save)));
  name.focus();
}

export function projectRow(p, { onDelete } = {}) {
  return h('li', { class: 'row' },
    h('a', { class: 'row-main', href: `#/projects/${p.id}` }, h('strong', {}, p.name), h('span', { class: 'muted' }, p.templateName ? `Template: ${p.templateName}` : 'No template')),
    h('span', { class: 'muted row-time' }, ago(p.updatedAt)),
    onDelete && h('button', { class: 'icon-btn danger', type: 'button', 'aria-label': `Delete ${p.name}`, onclick: onDelete }, icon('trash')));
}

export async function deleteProject(ctx, p) {
  if (!(await confirmDialog({ title: 'Delete project?', body: `“${p.name}” will be permanently deleted. Templates are not affected.`, confirmLabel: 'Delete project', danger: true }))) return false;
  try { await ctx.store.projects.remove(p.id); ctx.toast('Project deleted.', 'ok'); return true; }
  catch (e) { console.error(e); ctx.toast(`Could not delete the project. ${e.message}`, 'error'); return false; }
}

export const newProjectEmpty = (ctx) => emptyState({ title: 'No projects yet', body: 'A project holds one presentation. Create one to get started.', actions: [h('button', { class: 'btn btn-primary', type: 'button', onclick: () => openNewProject(ctx) }, icon('plus'), 'New project')] });

export function projects(ctx, root) {
  const body = h('div');
  root.append(pageHeader({ title: 'Projects', actions: [h('button', { class: 'btn btn-primary', type: 'button', onclick: () => openNewProject(ctx) }, icon('plus'), 'New project')] }), body);
  const reload = loadInto(body, () => ctx.store.projects.list(), (list) => {
    if (!list.length) return body.append(newProjectEmpty(ctx));
    body.append(h('ul', { class: 'rows' }, list.map((p) => projectRow(p, { onDelete: async () => { if (await deleteProject(ctx, p)) reload(); } }))));
  });
}
