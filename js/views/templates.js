import { h, icon } from '../lib/dom.js';
import { ago, fmtBytes, MODE_LABEL, plural } from '../lib/format.js';
import { classSummary, fontLine, paletteRow } from '../ui/design-panel.js';
import { confirmDialog } from '../ui/modal.js';
import { emptyState, loadInto, pageHeader } from '../ui/states.js';
import { openPreview } from '../ui/preview.js';
import { openUpload, pickAndUpload } from '../ui/upload.js';

export function templateCard(ctx, t, { onDeleted, onChanged } = {}) {
  const ratio = t.width / t.height;
  const frame = h('div', { class: 'frame' }, h('span', {}, t.slideTitles?.[0] || t.name));
  frame.style.aspectRatio = `${t.width} / ${t.height}`;
  frame.style[ratio >= 4 / 3 ? 'width' : 'height'] = '84%';

  const remove = async () => {
    let used = 0;
    try { used = (await ctx.store.projects.list()).filter((p) => p.templateId === t.id).length; } catch { /* count is informational */ }
    const note = used ? ` ${plural(used, 'project')} use it and will keep their data, but lose the link to the file.` : '';
    if (!(await confirmDialog({ title: 'Delete template?', body: `“${t.name}” and its file will be permanently deleted.${note}`, confirmLabel: 'Delete template', danger: true }))) return;
    try { await ctx.store.templates.remove(t); ctx.toast('Template deleted.', 'ok'); onDeleted?.(); }
    catch (e) { console.error(e); ctx.toast(`Could not delete the template. ${e.message}`, 'error'); }
  };

  return h('article', { class: 'tpl' },
    h('button', { class: 'tpl-cover', type: 'button', 'aria-label': `Preview ${t.name}`, onclick: () => openPreview(ctx, t, { onChange: onChanged }) }, frame),
    h('div', { class: 'tpl-info' },
      h('h3', { title: t.name }, t.name),
      h('p', { class: 'muted' }, `${plural(t.slideCount, 'slide')}, ${fmtBytes(t.size)}`),
      h('p', { class: 'muted' }, `Added ${ago(t.createdAt)}`),
      t.design
        ? h('div', { class: 'tpl-design' }, paletteRow(t.design), h('p', { class: 'muted' }, fontLine(t.design)), h('p', { class: 'muted' }, classSummary(t.design)), h('p', { class: 'muted' }, MODE_LABEL[t.mode || 'preserve']))
        : h('p', { class: 'muted' }, 'Design not analyzed yet. Open Preview to analyze it.')),
    h('div', { class: 'tpl-actions' },
      h('button', { class: 'btn btn-quiet', type: 'button', onclick: () => openPreview(ctx, t, { onChange: onChanged }) }, icon('eye'), 'Preview'),
      h('button', { class: 'icon-btn danger', type: 'button', 'aria-label': `Delete ${t.name}`, onclick: remove }, icon('trash'))));
}

export function templates(ctx, root) {
  const body = h('div');
  const dz = h('div', { class: 'dropzone', tabindex: '0', role: 'button', 'aria-label': 'Upload a PPTX template' },
    icon('upload'), h('div', {}, h('strong', {}, 'Drop a .pptx file here'), h('p', { class: 'muted' }, 'or click to choose one. Files stay private to your account.')));
  const open = () => pickAndUpload(ctx, () => reload());
  dz.addEventListener('click', open);
  dz.addEventListener('keydown', (e) => { if (e.key === 'Enter' || e.key === ' ') { e.preventDefault(); open(); } });
  dz.addEventListener('dragover', (e) => { e.preventDefault(); dz.classList.add('over'); });
  dz.addEventListener('dragleave', () => dz.classList.remove('over'));
  dz.addEventListener('drop', (e) => { e.preventDefault(); dz.classList.remove('over'); const f = e.dataTransfer.files[0]; if (f) openUpload(ctx, f, () => reload()); });

  root.append(pageHeader({ title: 'Templates', sub: 'PPTX files SlideFlow will build presentations from.' }), dz, body);
  const reload = loadInto(body, () => ctx.store.templates.list(), (list) => {
    if (!list.length) return body.append(emptyState({ title: 'No templates yet', body: 'Upload a .pptx file to start your library. SlideFlow checks it and reads its slides before saving.' }));
    body.append(h('div', { class: 'tpl-grid' }, list.map((t) => templateCard(ctx, t, { onDeleted: () => reload(), onChanged: () => reload() }))));
  });
}
