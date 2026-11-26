// Style bar of the "Rendered" tab (Phase 3.2.6-D): the controls that change the style of the selected element in the PREVIEW: text colour, size, font, bold, italic, fill, outline colour, and a reset.
// Every change goes through the controller's `style(id, patch)` (Phase 3.2.6-C), which commits to the temporary edit-state and draws it; this file only reads the edit-state (to show what is there) and
// calls the controller. It never touches the slide's DOM, the scene or the OOXML; nothing is saved or exported (kept while the deck is open, like moves, resizes and text edits).
// The bar stays in place with nothing selected (its controls are disabled), so the layout does not jump. Controls that cannot apply to the selected element are disabled: text controls need drawn text, the text colour
// needs a run that is not a hyperlink, a connector line has no fill, a picture fill cannot be recoloured, the outline needs a solid outline.
// Keys typed in the bar stay in the bar: they must not reach the dialog's arrow-key slide navigation. CSP: no inline style attributes anywhere (the rules are in css/styles.css).
import { h } from '../lib/dom.js';

const FONTS = ['Arial', 'Arial Narrow', 'Calibri', 'Cambria', 'Consolas', 'Courier New', 'Georgia', 'Helvetica', 'Segoe UI', 'Tahoma', 'Times New Roman', 'Verdana']; // the families the renderer can draw with a self-hosted equivalent (render/fonts.js); any other name is accepted and drawn when the browser has it
const HINT = 'Select a shape or line to change its colours, text size, font, bold or italic. Changes are kept in the preview only.';
const DONE = 'Style changed in the preview (kept while this deck is open; nothing is saved).';
const REASON = {
  'not-styleable': 'This element cannot be styled.',
  'no-text': 'This shape has no text to style.',
  'no-color-target': 'Text that is only a link keeps the link colour.',
  'no-fill-target': 'A line has no fill.',
  'picture-fill': 'A picture fill cannot be recoloured here.',
  'no-outline': 'This element has no solid outline to recolour.',
  invalid: 'That value is not valid.',
  disposed: 'This slide is being drawn again. Try again in a moment.',
  'no-edit-state': 'Styles cannot be kept for this slide.',
};
let uid = 0;

/** The runs a text style reaches: every run, and the default run of a paragraph that has none (the same rule as the edit-state's). */
const targets = (t) => { const out = []; for (const p of t.paras) { let any = false; for (const r of p.runs) if (!r.br) { any = true; out.push(r); } if (!any && p.defRun) out.push(p.defRun); } return out; };
/** { value, mixed } of a list: the value every item has, or mixed when they differ (or there is none). */
const common = (list, eq = (a, b) => a === b) => (list.length && list.every((v) => v != null && eq(v, list[0])) ? { value: list[0], mixed: false } : { value: null, mixed: list.length > 0 });
const hexOf = (v) => (typeof v === 'string' && /^[0-9a-f]{6}$/i.test(v) ? `#${v.toLowerCase()}` : null);

/**
 * Returns { el, bind(binding|null), sync(), dispose() }. `edits` = the deck's temporary edit-state. bind({ sel, scene, onApplied }) connects the bar to the controller of the drawn slide
 * (`sel` = what attachSelection returned, `scene` = the scene it draws; `onApplied(result)` is called after every style change, e.g. to refresh the slide's notes); bind(null) when the slide goes.
 * sync() shows the state of the selected element; the host calls it whenever the selection changes.
 */
export function styleBar({ edits }) {
  let sel = null; let scene = null; let onApplied = null; let lastId = null; let gen = 0;
  const n = ++uid; const dl = `sf-sb-fonts-${n}`;
  const color = (ctl, label) => h('input', { type: 'color', class: 'sf-sb-color', 'data-ctl': ctl, 'aria-label': label, title: label });
  const field = (label, ...kids) => h('label', { class: 'sf-sb-f' }, h('span', {}, label), ...kids);
  const toggle = (ctl, label, face, cls) => h('button', { type: 'button', class: `sf-sb-tgl ${cls}`, 'data-ctl': ctl, 'aria-pressed': 'false', 'aria-label': label, title: label }, face);
  const c = {
    textColor: color('text-color', 'Text colour'),
    size: h('input', { type: 'number', class: 'sf-sb-num', 'data-ctl': 'size', min: 1, max: 4000, step: 0.5, 'aria-label': 'Font size in points', title: 'Font size (pt)', inputmode: 'decimal' }),
    font: h('input', { type: 'text', class: 'sf-sb-font', 'data-ctl': 'font', list: dl, autocomplete: 'off', spellcheck: 'false', maxlength: 100, 'aria-label': 'Font family', title: 'Font family' }),
    bold: toggle('bold', 'Bold', 'B', 'sf-sb-b'),
    italic: toggle('italic', 'Italic', 'I', 'sf-sb-i'),
    fill: color('fill', 'Fill colour'),
    line: color('line', 'Outline colour'),
    reset: h('button', { type: 'button', class: 'btn btn-quiet sf-sb-reset', 'data-ctl': 'reset' }, 'Reset style'),
  };
  const status = h('p', { class: 'muted sf-sb-status', role: 'status', 'aria-live': 'polite' }, HINT);
  const list = h('datalist', { id: dl }, FONTS.map((f) => h('option', { value: f })));
  const bar = h('div', { class: 'sf-style-bar', role: 'toolbar', 'aria-label': 'Style of the selected element', 'data-style-bar': 'true' },
    h('div', { class: 'sf-sb-g', role: 'group', 'aria-label': 'Text' }, field('Text', c.textColor), field('Size', c.size, h('span', { class: 'sf-sb-unit' }, 'pt')), field('Font', c.font), c.bold, c.italic),
    h('div', { class: 'sf-sb-g', role: 'group', 'aria-label': 'Shape' }, field('Fill', c.fill), field('Outline', c.line)),
    c.reset);
  const el = h('div', { class: 'sf-sb', 'data-sf-sb': 'true' }, bar, status, list);
  const say = (t) => { status.textContent = t; };

  /** What the bar needs to know about an element: where it can apply and the values it shows. null when nothing styleable is selected. */
  function describe(e) {
    const ok = e && e.origin === 'slide' && ((e.tag === 'sp' && e.kind === 'shape') || (e.tag === 'cxnSp' && e.kind === 'line'));
    if (!ok) return null;
    const st = edits.effectiveStyle(e) || {};
    const t = e.kind === 'shape' && e.text && e.text.paras ? edits.effectiveStyledText(e) : null;
    const drawn = !!(t && t.paras.some((p) => p.runs.length)); // (the renderer draws a text box only for a paragraph that has runs)
    const rs = drawn ? targets(t) : []; const fs = t && t.body && t.body.fontScale > 0 ? t.body.fontScale : 1;
    return {
      text: drawn, textColor: drawn && rs.some((r) => !r.link),
      color: common(rs.filter((r) => !r.link).map((r) => hexOf(r.hex))),
      size: common(rs.map((r) => (typeof r.sz === 'number' ? Math.round((r.sz / fs) * 100) / 100 : null)), (a, b) => Math.abs(a - b) < 0.1),
      font: common(rs.map((r) => (r.font ? String(r.font) : null)), (a, b) => a.toLowerCase() === b.toLowerCase()),
      b: common(rs.map((r) => !!r.b)), i: common(rs.map((r) => !!r.i)),
      fill: e.kind === 'shape' && st.fill && st.fill.type !== 'image' ? { solid: st.fill.type === 'solid' ? hexOf(st.fill.hex) : null, type: st.fill.type || 'none' } : null,
      line: st.line && st.line.type === 'solid' ? hexOf(st.line.hex) : null,
    };
  }

  /** Show the state of the selected element. A control being typed in is left alone unless `force`. */
  function update(force = false) {
    const id = sel ? sel.id : null; const e = id && scene ? scene.byId.get(id) : null; const d = e ? describe(e) : null;
    if (id !== lastId) { lastId = id; say(!id ? HINT : d ? '' : 'This element cannot be styled here.'); }
    const put = (input, v) => { if (force || input.ownerDocument.activeElement !== input) input.value = v; };
    const on = (input, flag) => { input.disabled = !flag; };
    on(c.textColor, d && d.textColor); on(c.size, d && d.text); on(c.font, d && d.text); on(c.bold, d && d.text); on(c.italic, d && d.text);
    on(c.fill, d && d.fill); on(c.line, d && d.line); on(c.reset, d && edits.hasStyle(id));
    const mark = (input, mixed, title) => { if (mixed) input.setAttribute('data-mixed', 'true'); else input.removeAttribute('data-mixed'); input.title = mixed ? `${title} (mixed)` : title; };
    put(c.textColor, (d && d.color.value) || '#000000'); mark(c.textColor, !!(d && d.color.mixed), 'Text colour');
    put(c.size, d && d.size.value != null ? String(d.size.value) : ''); c.size.placeholder = d && d.size.mixed ? 'Mixed' : '';
    put(c.font, d && d.font.value ? d.font.value : ''); c.font.placeholder = d && d.font.mixed ? 'Mixed' : d && d.text ? 'Theme font' : '';
    for (const [btn, v, label] of [[c.bold, d && d.b, 'Bold'], [c.italic, d && d.i, 'Italic']]) { btn.setAttribute('aria-pressed', v && !v.mixed && v.value ? 'true' : v && v.mixed ? 'mixed' : 'false'); btn.title = v && v.mixed ? `${label} (mixed)` : label; }
    put(c.fill, (d && d.fill && d.fill.solid) || '#ffffff'); mark(c.fill, !!(d && d.fill && !d.fill.solid), d && d.fill && d.fill.type !== 'solid' ? `Fill colour (now: ${d.fill.type === 'none' ? 'no fill' : d.fill.type})` : 'Fill colour');
    put(c.line, (d && d.line) || '#000000');
  }

  /** Apply a patch to the selected element through the controller. Resolves to the controller's result, or null when nothing is selected. */
  async function apply(patch, force = true) {
    const id = sel ? sel.id : null; const my = gen; if (id == null) return null;
    let r; try { r = await sel.style(id, patch); } catch { r = { ok: false, reason: 'not-styleable' }; }
    if (my !== gen) return r; // the slide was drawn again meanwhile: the new one shows the state
    if (!r.ok) say(REASON[r.reason] || 'That style could not be applied.');
    else if (r.stale) { /* a newer change of this element is on its way and will say it */ }
    else if (r.changed && !r.drawn) say('Kept in memory, but this element could not be drawn with it.');
    else if (r.changed) say(DONE);
    update(force); // (a colour box being dragged in keeps what the person is doing: only the boxes that were typed in are put to their final value)
    if (r.ok && r.changed) onApplied?.(r);
    return r;
  }
  const hex = (input) => input.value.replace('#', '').toUpperCase();
  c.textColor.addEventListener('input', () => apply({ text: { color: hex(c.textColor) } }, false));
  c.fill.addEventListener('input', () => apply({ fill: { hex: hex(c.fill) } }, false));
  c.line.addEventListener('input', () => apply({ line: { hex: hex(c.line) } }, false));
  c.size.addEventListener('change', () => {
    const t = c.size.value.trim(); const v = Number(t);
    if (t === '') return apply({ text: { size: null } }); // an empty box goes back to the file's size
    if (!Number.isFinite(v) || v < 1 || v > 4000) { say('Size must be between 1 and 4000 pt.'); update(true); return null; }
    return apply({ text: { size: v } });
  });
  c.font.addEventListener('change', () => { const t = c.font.value.trim(); return apply({ text: { font: t === '' ? null : t } }); }); // an empty box goes back to the file's font
  for (const [btn, key] of [[c.bold, 'b'], [c.italic, 'i']]) btn.addEventListener('click', () => { const on = btn.getAttribute('aria-pressed') === 'true'; return apply({ text: { [key]: !on } }); });
  c.reset.addEventListener('click', () => apply({ text: null, fill: null, line: null }));
  // keys typed here are the bar's: the dialog turns arrow keys into slide changes; Escape in a box puts its value back instead of closing the dialog
  bar.addEventListener('keydown', (e) => {
    e.stopPropagation();
    if (e.key === 'Escape' && (e.target === c.size || e.target === c.font)) { e.preventDefault(); update(true); }
  });
  update();

  return {
    el,
    bind(b) { gen++; sel = b ? b.sel : null; scene = b ? b.scene : null; onApplied = b ? b.onApplied : null; lastId = undefined; update(true); },
    sync() { update(); },
    dispose() { gen++; sel = null; scene = null; onApplied = null; },
  };
}
