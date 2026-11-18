// Renders a stored Design System (js/pptx/design.js) as readable UI. Pure view code: no analysis happens here.
import { h } from '../lib/dom.js';
import { CLASS_LABEL, MODE_HELP, MODE_LABEL, plural, sizeLabel } from '../lib/format.js';

export function swatch(hex, title) {
  const s = h('span', { class: 'sw', title: title || `#${hex}` });
  s.style.background = `#${hex}`;
  return s;
}
/** Up to 5 swatches: used accent colours first, then the theme's own accents (hover shows the slot), then a neutral. */
export function paletteRow(design) {
  const used = design?.palette?.colors || [];
  const theme = design?.theme?.colors || {};
  const list = used.filter((x) => !x.neutral).slice(0, 4).map((x) => [x.hex, `#${x.hex}`]);
  for (const k of ['accent1', 'accent2', 'accent3', 'accent4', 'accent5', 'accent6', 'dk2']) {
    if (list.length >= 5) break;
    if (theme[k] && !list.some(([hex]) => hex === theme[k])) list.push([theme[k], `${k} #${theme[k]} (theme)`]);
  }
  if (list.length < 5) used.filter((x) => x.neutral).slice(0, 5 - list.length).forEach((x) => list.push([x.hex, `#${x.hex}`]));
  return h('span', { class: 'sw-row' }, list.map(([hex, title]) => swatch(hex, title)));
}
export const fontLine = (d) => [d?.modes?.preserve?.tokens?.fonts?.heading, d?.modes?.preserve?.tokens?.fonts?.body].filter((v, i, a) => v && a.indexOf(v) === i).join(' / ') || 'No text fonts found';
export const classSummary = (d) => Object.entries(d?.classification?.counts || {}).filter(([, n]) => n).slice(0, 4).map(([k, n]) => `${n} ${CLASS_LABEL[k].toLowerCase()}`).join(', ');

const THEME_ORDER = ['dk1', 'lt1', 'dk2', 'lt2', 'accent1', 'accent2', 'accent3', 'accent4', 'accent5', 'accent6'];
const sec = (title, ...body) => h('section', { class: 'dsec' }, h('h4', {}, title), ...body);
const table = (head, rows) => h('table', { class: 'dtable' }, h('thead', {}, h('tr', {}, head.map((x) => h('th', {}, x)))), h('tbody', {}, rows.map((r) => h('tr', {}, r.map((c) => h('td', {}, c))))));
const dl = (pairs) => h('dl', { class: 'facts' }, pairs.flatMap(([k, v]) => [h('dt', {}, k), h('dd', {}, v)]));
function bar(share) {
  const i = h('i'); i.style.width = `${Math.max(2, Math.round(share * 100))}%`;
  return h('span', { class: 'bar' }, h('span', { class: 'bar-track' }, i), `${Math.round(share * 100)}%`);
}
const pc = (v) => (v == null ? 'not available' : `${v}%`);
const pt = (v) => (v == null ? 'n/a' : `${v} pt`);

export function designPanel(t, d, { onMode, onReanalyze }) {
  const { preserve, enhance } = d.modes;
  const mode = t.mode || 'preserve';
  const tok = (m, f) => f(m.tokens);
  const cmp = [
    ['Heading font', (x) => x.fonts.heading || 'n/a'], ['Body font', (x) => x.fonts.body || 'n/a'],
    ['Title size', (x) => pt(x.sizesPt.title)], ['Body size', (x) => pt(x.sizesPt.body)], ['Footer size', (x) => pt(x.sizesPt.footer)],
    ['Margins L / R / T / B', (x) => ['left', 'right', 'top', 'bottom'].map((k) => x.marginsPct[k] ?? '–').join(' / ') + ' %'],
    ['Body text colour', (x) => (x.colors.text ? h('span', { class: 'inl' }, swatch(x.colors.text), `#${x.colors.text}`) : 'n/a')],
  ];
  const text = (n) => (n instanceof Node ? n.textContent : n);
  const modeOpt = (v) => h('label', { class: 'mode-opt' }, h('input', { type: 'radio', name: 'tpl-mode', value: v, checked: mode === v, onchange: () => onMode(v) }), h('span', {}, h('strong', {}, MODE_LABEL[v]), h('small', { class: 'muted' }, MODE_HELP[v])));

  const colors = d.palette.colors;
  const tp = d.typography;
  const m = d.spacing;
  const roleRows = ['title', 'subtitle', 'body', 'footer'].filter((r) => tp.roles[r]).map((r) => [r[0].toUpperCase() + r.slice(1), tp.roles[r].font || 'unknown', pt(tp.roles[r].sizePt), tp.roles[r].bold ? 'Bold' : 'Regular', h('span', { class: 'inl' }, swatch(tp.roles[r].color), `#${tp.roles[r].color}`)]);

  return h('div', { class: 'design' },
    h('fieldset', { class: 'modes' }, h('legend', {}, 'Template mode'), modeOpt('preserve'), modeOpt('enhance')),
    sec('What each mode would use',
      table(['', 'Preserve Design', 'Smart Enhance'], cmp.map(([k, f]) => {
        const a = tok(preserve, f); const b = tok(enhance, f);
        return [k, a, h('span', { class: text(a) !== text(b) ? 'diff' : '' }, b)];
      })),
      h('p', { class: 'muted fine' }, d.modes.note)),
    enhance.findings.length ? sec('Readability findings', h('ul', { class: 'finds' }, enhance.findings.map((f) => h('li', {}, h('span', { class: `tag ${f.severity}` }, f.severity === 'warn' ? 'Fix' : 'Note'), f.message, f.slides.length ? h('span', { class: 'muted' }, ` Slides: ${f.slides.join(', ')}.`) : null, f.evidence?.length ? h('span', { class: 'muted' }, ` ${f.evidence.map((e) => (e.ratio != null ? `Slide ${e.slide}: #${e.text} on #${e.background}, ${e.ratio}:1${e.sample ? ` ("${e.sample}")` : ''}` : `Slide ${e.slide}: ${e.kind} past the ${(e.edges || []).join(', ') || 'slide'} edge${e.byPct ? ` by ${e.byPct}%` : ''}${e.sample ? ` ("${e.sample}")` : ''}`)).join('; ')}.`) : null))))
      : sec('Readability findings', h('p', { class: 'muted' }, 'No readability problems were found by the checks Smart Enhance runs.')),
    enhance.changes.length > 0 && sec('Smart Enhance changes', h('ul', { class: 'finds' }, enhance.changes.map((c) => h('li', {}, h('strong', {}, c.token), `: ${c.from} to ${c.to}. `, h('span', { class: 'muted' }, c.reason))))),

    sec('Colours',
      d.theme.colors && Object.keys(d.theme.colors).length ? h('div', {}, h('p', { class: 'muted fine' }, `Theme${d.theme.name ? `: ${d.theme.name}` : ''}`), h('span', { class: 'sw-row big' }, THEME_ORDER.filter((k) => d.theme.colors[k]).map((k) => swatch(d.theme.colors[k], `${k} #${d.theme.colors[k]}`)))) : h('p', { class: 'muted' }, 'No theme colours available.'),
      h('p', { class: 'muted fine' }, 'Colours in use (share of text, fills and backgrounds)'),
      table(['Colour', 'Theme slot', 'Used for', 'Weight'], colors.slice(0, 8).map((c) => [h('span', { class: 'inl' }, swatch(c.hex), `#${c.hex}`), c.slot || 'custom', c.roles.join(', '), `${Math.round(c.weight * 100)}%`])),
      d.palette.offThemeCount ? h('p', { class: 'muted fine' }, `${plural(d.palette.offThemeCount, 'custom colour')} outside the theme palette.`) : null),

    sec('Typography',
      table(['Font', 'Share of text'], tp.fonts.map((f) => [f.name, bar(f.share)])),
      roleRows.length ? table(['Role', 'Font', 'Size', 'Weight', 'Colour'], roleRows) : h('p', { class: 'muted' }, 'No text styles found.'),
      tp.sizes.length ? h('p', { class: 'muted fine' }, `Sizes in use: ${tp.sizes.map((s) => s.pt).join(', ')} pt. Bold covers ${Math.round(tp.boldShare * 100)}% of text.`) : null),

    sec('Layout and spacing', dl([
      ['Slide', `${sizeLabel(d.slide.width, d.slide.height)}`],
      ['Margins (median)', `left ${pc(m.marginsPct.left)}, right ${pc(m.marginsPct.right)}, top ${pc(m.marginsPct.top)}, bottom ${pc(m.marginsPct.bottom)}`],
      ['Margin variation', `left ${pc(m.spreadPct.left)}, top ${pc(m.spreadPct.top)} (10th to 90th percentile)`],
      ['Gap between stacked items', pc(m.gapPct)],
      m.bleedSlides ? ['Decorative bleed', `${plural(m.bleedSlides, 'slide')} ${m.bleedSlides === 1 ? 'has' : 'have'} pictures or shapes placed flush with or past the edge on purpose; they are left out of the margins and are not a problem.${(() => { const t = d.slides.filter((s) => s.bleedDetail?.some((b) => b.repeats)).length; return t ? ` ${plural(t, 'slide')} repeat${t === 1 ? 's' : ''} the same shape at the same place, so it is part of the template.` : ''; })()}`] : null,
      m.offSlideSlides ? ['Content overflow', `${plural(m.offSlideSlides, 'slide')} ${m.offSlideSlides === 1 ? 'has' : 'have'} text, a table, a chart or a diagram running past the slide edge. See the readability findings for where.`] : null,
      ['Text alignment', Object.entries(m.alignment).map(([k, v]) => `${k} ${Math.round(v * 100)}%`).join(', ') || 'n/a'],
      ['Text insets', `${m.insetsSource}`],
      ['Backgrounds', Object.entries(d.backgrounds.types).map(([k, v]) => `${k} ${v}`).join(', ')],
      ['Shapes', Object.entries(d.shapes.byKind).map(([k, v]) => `${k} ${v}`).join(', ') || 'none'],
    ].filter(Boolean))),

    sec('Slide types',
      h('div', { class: 'chips' }, Object.entries(d.classification.counts).filter(([, n]) => n).map(([k, n]) => h('span', { class: 'chip' }, `${CLASS_LABEL[k]} ${n}`))),
      h('p', { class: 'muted fine' }, `Confidence: ${d.classification.confidence.high} high, ${d.classification.confidence.medium} medium, ${d.classification.confidence.low} low. Open the Slides tab to see each slide's type and the reasons.`)),

    sec('Layouts and patterns',
      table(['Layout', 'Type', 'Used by', 'Fixed shapes'], d.layouts.list.map((l) => [l.name || 'Unnamed', l.type || 'custom', plural(l.used, 'slide'), String(l.decor)])),
      d.patterns.length ? h('ul', { class: 'finds' }, d.patterns.map((p) => h('li', {}, `Repeated structure${p.layout ? ` on "${p.layout}"` : ''}: slides ${p.slides.join(', ')} (${CLASS_LABEL[p.class] || p.class}).`))) : h('p', { class: 'muted' }, 'No repeated slide structures found.'),
      d.recurring.length ? h('ul', { class: 'finds' }, d.recurring.map((r) => h('li', {}, `${r.kind === 'image' ? 'Picture' : 'Text'} "${r.label}" repeats on ${r.count} slides at the same position.`))) : null,
      h('p', { class: 'muted fine' }, `${plural(d.layouts.masters.length, 'slide master')}. Fixed shapes are logos, bars and frames defined on layouts or masters.`)),

    sec('Limits of this analysis', h('ul', { class: 'finds' }, d.limitations.map((l) => h('li', {}, l)))),
    h('div', { class: 'actions-left' }, h('button', { class: 'btn', type: 'button', onclick: onReanalyze }, 'Re-analyze'), h('button', { class: 'btn', type: 'button', onclick: () => download(t, d) }, 'Download design JSON')));
}

function download(t, d) {
  const url = URL.createObjectURL(new Blob([JSON.stringify(d, null, 2)], { type: 'application/json' }));
  const a = h('a', { href: url, download: `${(t.name || 'template').replace(/[^\w-]+/g, '-')}.design.json` });
  document.body.append(a); a.click(); a.remove();
  setTimeout(() => URL.revokeObjectURL(url), 1000);
}
