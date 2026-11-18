import { h } from '../lib/dom.js';
import { ago, fmtBytes, MODE_LABEL, plural, sizeLabel } from '../lib/format.js';
import { emptyState, loadInto, pageHeader } from '../ui/states.js';
import { openPreview } from '../ui/preview.js';
import { deleteProject } from './projects.js';

// Phase 1: a project is a record that points at a template. The slide editor and AI generation arrive in later phases.
export function project(ctx, root, id) {
  const body = h('div');
  root.append(h('a', { class: 'back', href: '#/projects' }, 'Projects'), body);
  loadInto(body, async () => {
    const p = await ctx.store.projects.get(id);
    const t = p?.templateId ? await ctx.store.templates.get(p.templateId).catch(() => null) : null;
    return { p, t };
  }, ({ p, t }) => {
    if (!p) return body.append(emptyState({ title: 'Project not found', body: 'It may have been deleted.', actions: [h('a', { class: 'btn', href: '#/projects' }, 'Back to projects')] }));
    document.title = `${p.name} · SlideFlow`;
    body.append(
      pageHeader({ title: p.name, sub: `Created ${ago(p.createdAt)}`, actions: [h('button', { class: 'btn btn-danger-quiet', type: 'button', onclick: async () => { if (await deleteProject(ctx, p)) location.hash = '#/projects'; } }, 'Delete project')] }),
      h('section', { class: 'panel' }, h('h2', {}, 'Template'), h('p', { class: 'muted' }, `Design mode: ${MODE_LABEL[p.mode || 'preserve']}`),
        t ? h('div', { class: 'split' },
          h('div', {}, h('strong', {}, t.name), h('p', { class: 'muted' }, `${plural(t.slideCount, 'slide')}, ${sizeLabel(t.width, t.height)}, ${fmtBytes(t.size)}`)),
          h('button', { class: 'btn', type: 'button', onclick: () => openPreview(ctx, t) }, 'Preview template'))
          : h('p', { class: 'muted' }, p.templateId ? `The template “${p.templateName}” has been deleted. This project's data is unchanged.` : 'This project does not use a template.')),
      h('section', { class: 'panel' }, h('h2', {}, 'Slides'), h('p', { class: 'muted' }, 'Slide generation and editing are not available yet. This project is saved and ready for them.')));
  });
}
