// Structural preview: draws text, picture and object boxes at their real positions from the parsed XML.
// It does NOT render fonts, colours, gradients, effects or charts, and says so in the UI.
import { h } from '../lib/dom.js';
import { plural, sizeLabel } from '../lib/format.js';
import { parsePptx, mediaUrl } from '../pptx/parser.js';
import { DESIGN_VERSION, extractDesign } from '../pptx/design.js';
import { CLASS_LABEL } from '../lib/format.js';
import { designPanel, swatch } from './design-panel.js';
import { modal } from './modal.js';
import { errorState, friendly, spinner } from './states.js';

const KIND_LABEL = { chart: 'Chart', object: 'Embedded object', table: 'Table' };

export async function openPreview(ctx, t, { onChange } = {}) {
  const m = modal({ title: t.name, wide: true });
  m.body.append(h('p', { class: 'status' }, spinner(), 'Opening template…'));
  const urls = new Map();
  let changed = false;
  m.onClose = () => { urls.forEach((u) => u && URL.revokeObjectURL(u)); if (changed) onChange?.(); };

  let parsed;
  try { parsed = await parsePptx(await ctx.store.templates.file(t)); }
  catch (e) {
    console.error(e);
    const msg = e.name === 'PptxError' ? e.message : `Couldn't load the template file. ${friendly(e)}${ctx.store.mode === 'firebase' ? ' If this keeps happening, check the Storage bucket CORS setup in the README.' : ''}`;
    return m.body.replaceChildren(errorState(new Error(msg)));
  }

  // Design: use the stored analysis when it matches the current engine, otherwise analyse now and cache it on the template.
  let design = t.design && t.designVersion === DESIGN_VERSION ? t.design : null;
  let designError = null;
  const analyze = async (explicit) => {
    designError = null;
    try {
      design = await extractDesign(parsed);
      try { Object.assign(t, await ctx.store.templates.update(t, { design, designVersion: design.version })); changed = true; if (explicit) ctx.toast('Design analysis saved.', 'ok'); }
      catch (e) { console.error(e); ctx.toast(`Analysis done but could not be saved. ${e.message}`, 'error'); }
    } catch (e) { console.error(e); design = null; designError = e; }
  };
  if (!design) await analyze(false);

  const { width: W, height: H } = parsed;
  const pct = (v, total) => `${(v / total) * 100}%`;
  const slide = h('div', { class: 'slide' });
  slide.style.aspectRatio = `${W} / ${H}`;
  const details = h('div', { class: 'details' });
  const items = parsed.slides.map((s, i) => h('li', {}, h('button', { type: 'button', class: 'film-item', 'data-i': i, onclick: () => select(i) },
    h('span', { class: 'film-n' }, i + 1), h('span', { class: 'film-t' }, s.title || 'No title'), design && h('span', { class: 'tag' }, CLASS_LABEL[design.slides[i]?.class] || ''), s.hidden && h('span', { class: 'tag' }, 'Hidden'))));
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

  function classInfo(i) {
    const c = design?.slides[i];
    if (!c) return [];
    return [
      h('h4', {}, 'Slide type'),
      h('p', {}, `${CLASS_LABEL[c.class]} `, h('span', { class: 'muted' }, `(${c.confidence} confidence)`)),
      h('ul', { class: 'file-list' }, c.reasons.map((r) => h('li', {}, r))),
      h('p', { class: 'muted' }, h('span', { class: 'inl' }, 'Background: ', c.background.hex ? swatch(c.background.hex) : null, `${c.background.type}${c.background.hex ? ` #${c.background.hex}` : ''}, from ${c.background.source}`)),
      c.minPt != null ? h('p', { class: 'muted' }, `Text sizes: ${c.minPt === c.maxPt ? c.minPt : `${c.minPt} to ${c.maxPt}`} pt`) : null,
    ];
  }

  function slideDetails(s, i) {
    const texts = s.shapes.filter((x) => x.text);
    const unplaced = s.shapes.filter((x) => !x.box).length;
    return [
      h('h3', {}, `Slide ${i + 1}`),
      h('p', { class: 'muted' }, s.layout ? `Layout: ${s.layout}` : 'Layout: not specified'),
      ...classInfo(i),
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

  const slidesTab = h('div', { class: 'preview' },
    h('ol', { class: 'film', 'aria-label': 'Slides' }, items),
    h('div', { class: 'stage-wrap' }, slide, h('p', { class: 'muted fine' }, 'Structural preview: text, pictures and positions only. Fonts, colours, effects and charts are not rendered.')),
    details);
  const designTab = h('div', { class: 'design-host' });
  const renderDesign = () => {
    designTab.replaceChildren(design
      ? designPanel(t, design, {
        onMode: async (mode) => {
          try { Object.assign(t, await ctx.store.templates.update(t, { mode })); changed = true; ctx.toast(`Mode set to ${mode === 'enhance' ? 'Smart Enhance' : 'Preserve Design'}.`, 'ok'); }
          catch (e) { console.error(e); ctx.toast(`Could not change the mode. ${e.message}`, 'error'); renderDesign(); }
        },
        onReanalyze: async () => { designTab.replaceChildren(h('p', { class: 'status' }, spinner(), 'Analyzing design…')); await analyze(true); renderDesign(); select(Math.max(current, 0)); },
      })
      : errorState(new Error(`Design analysis failed: ${designError?.message || 'unknown error'}`), async () => { designTab.replaceChildren(h('p', { class: 'status' }, spinner(), 'Analyzing design…')); await analyze(true); renderDesign(); }));
  };
  renderDesign();
  designTab.hidden = true;
  const tabs = [['Slides', slidesTab], ['Design', designTab]].map(([label, panel], k) => h('button', { type: 'button', role: 'tab', class: 'tab', 'aria-selected': String(k === 0), onclick: (e) => {
    tabs.forEach((b, j) => b.setAttribute('aria-selected', String(b === e.currentTarget)));
    [slidesTab, designTab].forEach((p) => { p.hidden = p !== panel; });
  } }, label));

  m.body.replaceChildren(...[
    h('p', { class: 'muted summary' }, `${plural(parsed.slideCount, 'slide')}, ${sizeLabel(W, H)}, ${plural(parsed.imageRefs, 'picture reference')}`),
    h('div', { class: 'tabs', role: 'tablist' }, tabs),
    slidesTab, designTab,
    parsed.warnings.length > 0 && h('details', { class: 'notice' }, h('summary', {}, plural(parsed.warnings.length, 'note')), h('ul', {}, parsed.warnings.map((w) => h('li', {}, w))))].filter(Boolean));
  select(0);
}
