// Structural preview: draws text, picture and object boxes at their real positions from the parsed XML.
// It does NOT render fonts, colours, gradients, effects or charts, and says so in the UI.
import { h } from '../lib/dom.js';
import { plural, sizeLabel } from '../lib/format.js';
import { parsePptx, mediaUrl } from '../pptx/parser.js';
import { modal } from './modal.js';
import { errorState, friendly, spinner } from './states.js';

const KIND_LABEL = { chart: 'Chart', object: 'Embedded object', table: 'Table' };

export async function openPreview(ctx, t) {
  const m = modal({ title: t.name, wide: true });
  m.body.append(h('p', { class: 'status' }, spinner(), 'Opening template…'));
  const urls = new Map();
  m.onClose = () => urls.forEach((u) => u && URL.revokeObjectURL(u));

  let parsed;
  try { parsed = await parsePptx(await ctx.store.templates.file(t)); }
  catch (e) {
    console.error(e);
    const msg = e.name === 'PptxError' ? e.message : `Couldn't load the template file. ${friendly(e)}${ctx.store.mode === 'firebase' ? ' If this keeps happening, check the Storage bucket CORS setup in the README.' : ''}`;
    return m.body.replaceChildren(errorState(new Error(msg)));
  }

  const { width: W, height: H } = parsed;
  const pct = (v, total) => `${(v / total) * 100}%`;
  const slide = h('div', { class: 'slide' });
  slide.style.aspectRatio = `${W} / ${H}`;
  const details = h('div', { class: 'details' });
  const items = parsed.slides.map((s, i) => h('li', {}, h('button', { type: 'button', class: 'film-item', 'data-i': i, onclick: () => select(i) },
    h('span', { class: 'film-n' }, i + 1), h('span', { class: 'film-t' }, s.title || 'No title'), s.hidden && h('span', { class: 'tag' }, 'Hidden'))));
  let current = -1;

  async function select(i) {
    current = i;
    const s = parsed.slides[i];
    items.forEach((li, k) => { const b = li.firstChild; b.setAttribute('aria-current', String(k === i)); if (k === i) b.scrollIntoView({ block: 'nearest', inline: 'nearest' }); });
    slide.replaceChildren();
    const pending = [];
    for (const sh of s.shapes) {
      if (!sh.box) continue;
      const el = h('div', { class: `shp ${sh.kind}${sh.placeholder === 'title' ? ' is-title' : ''}`, title: sh.name });
      Object.assign(el.style, { left: pct(sh.box.x, W), top: pct(sh.box.y, H), width: pct(sh.box.w, W), height: pct(sh.box.h, H) });
      if (sh.text) el.textContent = sh.text.replace(/\t/g, '  ');
      else if (sh.kind in KIND_LABEL) el.textContent = KIND_LABEL[sh.kind];
      if (sh.kind === 'image' && sh.image && !sh.image.external) pending.push([el, sh.image.target]);
      else if (sh.kind === 'image') el.textContent = 'Picture not available';
      slide.append(el);
    }
    details.replaceChildren(...slideDetails(s, i).filter(Boolean));
    for (const [el, target] of pending) {
      if (!urls.has(target)) urls.set(target, await mediaUrl(parsed.zip, target).catch(() => null));
      if (current !== i) return;
      const url = urls.get(target);
      if (url) el.append(h('img', { src: url, alt: '' }));
      else el.textContent = 'Picture format not previewable';
    }
  }

  function slideDetails(s, i) {
    const texts = s.shapes.filter((x) => x.text);
    const unplaced = s.shapes.filter((x) => !x.box).length;
    return [
      h('h3', {}, `Slide ${i + 1}`),
      h('p', { class: 'muted' }, s.layout ? `Layout: ${s.layout}` : 'Layout: not specified'),
      h('h4', {}, 'Text'),
      texts.length ? h('ul', { class: 'text-list' }, texts.map((x) => h('li', {}, x.placeholder && h('span', { class: 'tag' }, x.placeholder), x.text))) : h('p', { class: 'muted' }, 'No text on this slide.'),
      h('h4', {}, 'Pictures'),
      s.images.length ? h('ul', { class: 'file-list' }, s.images.map((im) => h('li', {}, im.external ? `Linked externally: ${im.target}` : im.target.replace(/^ppt\//, '')))) : h('p', { class: 'muted' }, 'No pictures referenced.'),
      unplaced > 0 && h('p', { class: 'muted' }, `${plural(unplaced, 'shape')} inherit their position from a layout that could not be resolved and are not drawn.`),
    ];
  }

  m.dlg.addEventListener('keydown', (e) => {
    const next = { ArrowDown: 1, ArrowRight: 1, ArrowUp: -1, ArrowLeft: -1 }[e.key];
    if (!next) return;
    const i = Math.min(Math.max(current + next, 0), items.length - 1);
    if (i !== current) { e.preventDefault(); select(i); items[i].firstChild.focus(); }
  });

  m.body.replaceChildren(...[
    h('p', { class: 'muted summary' }, `${plural(parsed.slideCount, 'slide')}, ${sizeLabel(W, H)}, ${plural(parsed.imageRefs, 'picture reference')}`),
    h('div', { class: 'preview' },
      h('ol', { class: 'film', 'aria-label': 'Slides' }, items),
      h('div', { class: 'stage-wrap' }, slide, h('p', { class: 'muted fine' }, 'Structural preview: text, pictures and positions only. Fonts, colours, effects and charts are not rendered.')),
      details),
    parsed.warnings.length > 0 && h('details', { class: 'notice' }, h('summary', {}, plural(parsed.warnings.length, 'note')), h('ul', {}, parsed.warnings.map((w) => h('li', {}, w))))].filter(Boolean));
  select(0);
}
