import { h, icon } from '../lib/dom.js';
import { emptyState, loadInto, pageHeader } from '../ui/states.js';
import { pickAndUpload } from '../ui/upload.js';
import { deleteProject, newProjectEmpty, openNewProject, projectRow } from './projects.js';
import { templateCard } from './templates.js';

export function dashboard(ctx, root) {
  const body = h('div');
  const upload = () => pickAndUpload(ctx, () => reload());
  root.append(pageHeader({ title: 'Dashboard', actions: [
    h('button', { class: 'btn btn-primary', type: 'button', onclick: () => openNewProject(ctx) }, icon('plus'), 'New project'),
    h('button', { class: 'btn', type: 'button', onclick: upload }, icon('upload'), 'Upload template')] }), body);

  const reload = loadInto(body, () => Promise.all([ctx.store.projects.list(), ctx.store.templates.list()]), ([projects, templates]) => {
    const section = (title, link, ...content) => h('section', { class: 'section' }, h('div', { class: 'section-head' }, h('h2', {}, title), link), ...content);
    body.append(
      section('Recent projects', projects.length > 5 && h('a', { href: '#/projects' }, 'View all'),
        projects.length ? h('ul', { class: 'rows' }, projects.slice(0, 5).map((p) => projectRow(p, { onDelete: async () => { if (await deleteProject(ctx, p)) reload(); } }))) : newProjectEmpty(ctx)),
      section('Template library', templates.length > 4 && h('a', { href: '#/templates' }, 'Manage templates'),
        templates.length ? h('div', { class: 'tpl-grid' }, templates.slice(0, 4).map((t) => templateCard(ctx, t, { onDeleted: () => reload() })))
          : emptyState({ title: 'No templates yet', body: 'Upload a .pptx file. SlideFlow validates it and reads its slides before saving.', actions: [h('button', { class: 'btn', type: 'button', onclick: upload }, icon('upload'), 'Upload template')] })));
  });
}
