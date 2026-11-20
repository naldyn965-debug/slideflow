// Shape geometry (Phase 3.1): turns a preset shape name (prstGeom) or a custom geometry (custGeom) into SVG path data.
// Everything is in points (EMU / 12700) inside the unrotated shape box, so the SVG viewBox is exactly the box and strokes are never distorted.
// Pure functions, no DOM reads except the custGeom element passed in. Each result: { paths:[{ d, fill, stroke }], textRect:[l,t,r,b]|null, approx:string|null }.
//   - Presets listed in PRESETS follow the ECMA-376 presetShapeDefinitions formulas (adjust handles honoured).
//   - Presets listed in SIMPLIFIED are drawn with a close, hand-made outline and are reported as approximated.
//   - Any other preset returns null: the renderer draws the bounding rectangle and reports the shape as approximated.
import { kid, kids } from '../pptx/parser.js';

const f = (n) => (Math.round(n * 1000) / 1000).toString();
const rad = (deg) => (deg * Math.PI) / 180;
const pin = (lo, v, hi) => Math.min(hi, Math.max(lo, v));

// ---------- arcs (OOXML arcTo: visual angles on an ellipse, clockwise, starting at the current point) ----------
function param(wR, hR, deg) { const t = rad(deg); return Math.atan2(Math.sin(t) / hR, Math.cos(t) / wR); }
/** Returns { cmd, x, y } for an arcTo from point (x0,y0); angles in degrees. */
export function arcTo(x0, y0, wR, hR, stDeg, swDeg) {
  if (wR <= 0 || hR <= 0 || !swDeg) return { cmd: '', x: x0, y: y0 };
  const t0 = param(wR, hR, stDeg); const cx = x0 - wR * Math.cos(t0); const cy = y0 - hR * Math.sin(t0);
  const at = (deg) => { const t = param(wR, hR, deg); return [cx + wR * Math.cos(t), cy + hR * Math.sin(t)]; };
  const sweep = swDeg > 0 ? 1 : 0;
  if (Math.abs(swDeg) >= 359.999) { const [mx, my] = at(stDeg + swDeg / 2); return { cmd: `A${f(wR)} ${f(hR)} 0 0 ${sweep} ${f(mx)} ${f(my)} A${f(wR)} ${f(hR)} 0 0 ${sweep} ${f(x0)} ${f(y0)}`, x: x0, y: y0 }; }
  const [x1, y1] = at(stDeg + swDeg);
  return { cmd: `A${f(wR)} ${f(hR)} 0 ${Math.abs(swDeg) > 180 ? 1 : 0} ${sweep} ${f(x1)} ${f(y1)}`, x: x1, y: y1 };
}
const ellipse = (cx, cy, rx, ry) => `M${f(cx - rx)} ${f(cy)}A${f(rx)} ${f(ry)} 0 1 1 ${f(cx + rx)} ${f(cy)}A${f(rx)} ${f(ry)} 0 1 1 ${f(cx - rx)} ${f(cy)}Z`;
const poly = (pts) => `M${pts.map((p) => `${f(p[0])} ${f(p[1])}`).join('L')}Z`;
function star(n, w, h, adj, hf = 100000, vf = 100000) {
  const rx = (w / 2) * (hf / 100000); const ry = (h / 2) * (vf / 100000); const cx = w / 2; const cy = (h / 2) * (vf / 100000);
  const k = adj / 50000; const pts = [];
  for (let i = 0; i < n; i++) {
    const a = rad(-90 + (360 / n) * i); const b = rad(-90 + (360 / n) * (i + 0.5));
    pts.push([cx + rx * Math.cos(a), cy + ry * Math.sin(a)], [cx + rx * k * Math.cos(b), cy + ry * k * Math.sin(b)]);
  }
  return poly(pts);
}
const adjOf = (adj, name, d) => (adj && adj[name] != null ? adj[name] : d);

// ---------- presets (ECMA-376 presetShapeDefinitions) ----------
// Each: (w, h, A) => { d, text?:[l,t,r,b] }; A(name, default) reads an adjust value.
const rect = (w, h) => `M0 0H${f(w)}V${f(h)}H0Z`;
const PRESETS = {
  rect: (w, h) => ({ d: rect(w, h) }),
  roundRect: (w, h, A) => { const ss = Math.min(w, h); const r = (ss * pin(0, A('adj', 16667), 50000)) / 100000; const i = r * 0.29289;
    return { d: `M0 ${f(r)}A${f(r)} ${f(r)} 0 0 1 ${f(r)} 0H${f(w - r)}A${f(r)} ${f(r)} 0 0 1 ${f(w)} ${f(r)}V${f(h - r)}A${f(r)} ${f(r)} 0 0 1 ${f(w - r)} ${f(h)}H${f(r)}A${f(r)} ${f(r)} 0 0 1 0 ${f(h - r)}Z`, text: [i, i, w - i, h - i] }; },
  ellipse: (w, h) => ({ d: ellipse(w / 2, h / 2, w / 2, h / 2), text: [w * 0.14645, h * 0.14645, w * 0.85355, h * 0.85355] }),
  triangle: (w, h, A) => { const x = (w * pin(0, A('adj', 50000), 100000)) / 100000; return { d: poly([[0, h], [x, 0], [w, h]]), text: [x / 2, h / 2, (x + w) / 2, h] }; },
  rtTriangle: (w, h) => ({ d: poly([[0, 0], [0, h], [w, h]]), text: [0, h / 2, w / 2, h] }),
  diamond: (w, h) => ({ d: poly([[0, h / 2], [w / 2, 0], [w, h / 2], [w / 2, h]]), text: [w / 4, h / 4, (w * 3) / 4, (h * 3) / 4] }),
  parallelogram: (w, h, A) => { const ss = Math.min(w, h); const x = (ss * pin(0, A('adj', 25000), (100000 * w) / ss)) / 100000; return { d: poly([[0, h], [x, 0], [w, 0], [w - x, h]]), text: [x / 2, 0, w - x / 2, h] }; },
  trapezoid: (w, h, A) => { const ss = Math.min(w, h); const x = (ss * pin(0, A('adj', 25000), (50000 * w) / ss)) / 100000; return { d: poly([[0, h], [x, 0], [w - x, 0], [w, h]]), text: [x / 3, h / 3, w - x / 3, h] }; },
  hexagon: (w, h, A) => { const ss = Math.min(w, h); const x = (ss * pin(0, A('adj', 25000), (50000 * w) / ss)) / 100000; return { d: poly([[0, h / 2], [x, 0], [w - x, 0], [w, h / 2], [w - x, h], [x, h]]) }; },
  octagon: (w, h, A) => { const ss = Math.min(w, h); const x = (ss * pin(0, A('adj', 29289), 50000)) / 100000; return { d: poly([[0, x], [x, 0], [w - x, 0], [w, x], [w, h - x], [w - x, h], [x, h], [0, h - x]]), text: [x / 2, x / 2, w - x / 2, h - x / 2] }; },
  homePlate: (w, h, A) => { const ss = Math.min(w, h); const dx = (ss * pin(0, A('adj', 50000), (100000 * w) / ss)) / 100000; return { d: poly([[0, 0], [w - dx, 0], [w, h / 2], [w - dx, h], [0, h]]), text: [0, 0, (w - dx / 2 + w - dx) / 2, h] }; },
  chevron: (w, h, A) => { const ss = Math.min(w, h); const x = (ss * pin(0, A('adj', 50000), (100000 * w) / ss)) / 100000; return { d: poly([[0, 0], [w - x, 0], [w, h / 2], [w - x, h], [0, h], [x, h / 2]]), text: [x, 0, w - x, h] }; },
  rightArrow: (w, h, A) => { const ss = Math.min(w, h); const a1 = pin(0, A('adj1', 50000), 100000); const dx = (ss * pin(0, A('adj2', 50000), (100000 * w) / ss)) / 100000; const x1 = w - dx; const dy = (h * a1) / 200000; return { d: poly([[0, h / 2 - dy], [x1, h / 2 - dy], [x1, 0], [w, h / 2], [x1, h], [x1, h / 2 + dy], [0, h / 2 + dy]]) }; },
  leftArrow: (w, h, A) => { const ss = Math.min(w, h); const a1 = pin(0, A('adj1', 50000), 100000); const x2 = (ss * pin(0, A('adj2', 50000), (100000 * w) / ss)) / 100000; const dy = (h * a1) / 200000; return { d: poly([[0, h / 2], [x2, 0], [x2, h / 2 - dy], [w, h / 2 - dy], [w, h / 2 + dy], [x2, h / 2 + dy], [x2, h]]) }; },
  upArrow: (w, h, A) => { const ss = Math.min(w, h); const a1 = pin(0, A('adj1', 50000), 100000); const y2 = (ss * pin(0, A('adj2', 50000), (100000 * h) / ss)) / 100000; const dx = (w * a1) / 200000; return { d: poly([[w / 2, 0], [w, y2], [w / 2 + dx, y2], [w / 2 + dx, h], [w / 2 - dx, h], [w / 2 - dx, y2], [0, y2]]) }; },
  downArrow: (w, h, A) => { const ss = Math.min(w, h); const a1 = pin(0, A('adj1', 50000), 100000); const dy2 = (ss * pin(0, A('adj2', 50000), (100000 * h) / ss)) / 100000; const y1 = h - dy2; const dx = (w * a1) / 200000; return { d: poly([[w / 2 - dx, 0], [w / 2 + dx, 0], [w / 2 + dx, y1], [w, y1], [w / 2, h], [0, y1], [w / 2 - dx, y1]]) }; },
  plus: (w, h, A) => { const ss = Math.min(w, h); const x = (ss * pin(0, A('adj', 25000), 50000)) / 100000; return { d: poly([[0, x], [x, x], [x, 0], [w - x, 0], [w - x, x], [w, x], [w, h - x], [w - x, h - x], [w - x, h], [x, h], [x, h - x], [0, h - x]]), text: [0, x, w, h - x] }; },
  donut: (w, h, A) => { const ss = Math.min(w, h); const dr = (ss * pin(0, A('adj', 25000), 50000)) / 100000; return { d: ellipse(w / 2, h / 2, w / 2, h / 2) + ellipse(w / 2, h / 2, w / 2 - dr, h / 2 - dr), evenodd: true, text: [w * 0.14645, h * 0.14645, w * 0.85355, h * 0.85355] }; },
  blockArc: (w, h, A) => {
    const ss = Math.min(w, h); const st = pin(0, A('adj1', 10800000), 21599999) / 60000; const en = pin(0, A('adj2', 0), 21599999) / 60000; const dr = (ss * pin(0, A('adj3', 25000), 50000)) / 100000;
    let sw = en - st; if (sw <= 0) sw += 360;
    const wR = w / 2; const hR = h / 2; const ox = (deg) => { const t = param(wR, hR, deg); return [wR + wR * Math.cos(t), hR + hR * Math.sin(t)]; };
    const iw = wR - dr; const ih = hR - dr; const ix = (deg) => { const t = param(iw, ih, deg); return [wR + iw * Math.cos(t), hR + ih * Math.sin(t)]; };
    const [sx, sy] = ox(st); const a1 = arcTo(sx, sy, wR, hR, st, sw); const [jx, jy] = ix(en); const a2 = arcTo(jx, jy, iw, ih, en, -sw);
    return { d: `M${f(sx)} ${f(sy)}${a1.cmd}L${f(jx)} ${f(jy)}${a2.cmd}Z` };
  },
  flowChartProcess: (w, h) => ({ d: rect(w, h) }),
  flowChartAlternateProcess: (w, h) => PRESETS.roundRect(w, h, (n, d) => (n === 'adj' ? 16667 : d)),
  flowChartDecision: (w, h) => PRESETS.diamond(w, h, () => 0),
  flowChartTerminator: (w, h) => { const rx = (w * 3475) / 21600; return { d: `M${f(rx)} 0H${f(w - rx)}A${f(rx)} ${f(h / 2)} 0 0 1 ${f(w - rx)} ${f(h)}H${f(rx)}A${f(rx)} ${f(h / 2)} 0 0 1 ${f(rx)} 0Z`, text: [w * 0.1, h * 0.1, w * 0.9, h * 0.9] }; },
  flowChartPredefinedProcess: (w, h) => ({ d: rect(w, h) + `M${f(w / 8)} 0V${f(h)}M${f((w * 7) / 8)} 0V${f(h)}` }),
  flowChartDocument: (w, h) => ({ d: `M0 0H${f(w)}V${f(h * 0.8)}C${f(w * 0.75)} ${f(h * 0.6)} ${f(w * 0.5)} ${f(h * 0.6)} ${f(w * 0.5)} ${f(h * 0.8)}C${f(w * 0.25)} ${f(h)} ${f(w * 0.1)} ${f(h * 0.95)} 0 ${f(h * 0.9)}Z`, approx: 'document outline is a close approximation' }),
  flowChartInputOutput: (w, h) => ({ d: poly([[w / 5, 0], [w, 0], [(w * 4) / 5, h], [0, h]]) }),
  flowChartConnector: (w, h) => ({ d: ellipse(w / 2, h / 2, w / 2, h / 2) }),
  star4: (w, h, A) => ({ d: star(4, w, h, A('adj', 12500)) }),
  star5: (w, h, A) => ({ d: star(5, w, h, A('adj', 19098), 105146, 110557) }),
  star6: (w, h, A) => ({ d: star(6, w, h, A('adj', 28868), 115470) }),
  star8: (w, h, A) => ({ d: star(8, w, h, A('adj', 37500)) }),
  star10: (w, h, A) => ({ d: star(10, w, h, A('adj', 42533), 105146) }),
  star12: (w, h, A) => ({ d: star(12, w, h, A('adj', 37500)) }),
  star16: (w, h, A) => ({ d: star(16, w, h, A('adj', 37500)) }),
  star24: (w, h, A) => ({ d: star(24, w, h, A('adj', 37500)) }),
  line: (w, h) => ({ d: `M0 0L${f(w)} ${f(h)}`, open: true }),
  straightConnector1: (w, h) => ({ d: `M0 0L${f(w)} ${f(h)}`, open: true }),
  bentConnector2: (w, h) => ({ d: `M0 0H${f(w)}V${f(h)}`, open: true }),
  bentConnector3: (w, h, A) => { const x = (w * A('adj1', 50000)) / 100000; return { d: `M0 0H${f(x)}V${f(h)}H${f(w)}`, open: true }; },
  bentConnector4: (w, h, A) => { const x = (w * A('adj1', 50000)) / 100000; const y = (h * A('adj2', 50000)) / 100000; return { d: `M0 0H${f(x)}V${f(y)}H${f(w)}V${f(h)}`, open: true }; },
  bentConnector5: (w, h, A) => { const x1 = (w * A('adj1', 50000)) / 100000; const y2 = (h * A('adj2', 50000)) / 100000; const x3 = (w * A('adj3', 50000)) / 100000; return { d: `M0 0H${f(x1)}V${f(y2)}H${f(x3)}V${f(h)}H${f(w)}`, open: true }; },
  curvedConnector2: (w, h) => ({ d: `M0 0C${f(w / 2)} 0 ${f(w)} ${f(h / 2)} ${f(w)} ${f(h)}`, open: true }),
  curvedConnector3: (w, h, A) => { const x2 = (w * A('adj1', 50000)) / 100000; return { d: `M0 0C${f(x2 / 2)} 0 ${f(x2)} ${f(h / 4)} ${f(x2)} ${f(h / 2)}C${f(x2)} ${f((h * 3) / 4)} ${f((w + x2) / 2)} ${f(h)} ${f(w)} ${f(h)}`, open: true }; },
};
PRESETS.pentagon = PRESETS.homePlate;
PRESETS.flowChartSummingJunction = (w, h) => ({ d: ellipse(w / 2, h / 2, w / 2, h / 2) + `M${f(w * 0.1464)} ${f(h * 0.1464)}L${f(w * 0.8536)} ${f(h * 0.8536)}M${f(w * 0.8536)} ${f(h * 0.1464)}L${f(w * 0.1464)} ${f(h * 0.8536)}` });
PRESETS.snip1Rect = (w, h, A) => { const ss = Math.min(w, h); const x = (ss * pin(0, A('adj', 16667), 50000)) / 100000; return { d: poly([[0, 0], [w - x, 0], [w, x], [w, h], [0, h]]) }; };
PRESETS.round2SameRect = (w, h, A) => { const ss = Math.min(w, h); const a1 = (ss * pin(0, A('adj1', 16667), 50000)) / 100000; const a2 = (ss * pin(0, A('adj2', 0), 50000)) / 100000;
  return { d: `M0 ${f(a1)}A${f(a1)} ${f(a1)} 0 0 1 ${f(a1)} 0H${f(w - a1)}A${f(a1)} ${f(a1)} 0 0 1 ${f(w)} ${f(a1)}V${f(h - a2)}${a2 ? `A${f(a2)} ${f(a2)} 0 0 1 ${f(w - a2)} ${f(h)}H${f(a2)}A${f(a2)} ${f(a2)} 0 0 1 0 ${f(h - a2)}` : `V${f(h)}H0`}Z` }; };

// Shapes whose real outline is complex: drawn with a close hand-made outline and reported as approximated.
const SIMPLIFIED = {
  heart: (w, h) => ({ d: `M${f(w / 2)} ${f(h)}C${f(w * 0.1)} ${f(h * 0.7)} 0 ${f(h * 0.45)} 0 ${f(h * 0.28)}C0 ${f(h * 0.1)} ${f(w * 0.15)} 0 ${f(w * 0.3)} 0C${f(w * 0.4)} 0 ${f(w * 0.47)} ${f(h * 0.06)} ${f(w / 2)} ${f(h * 0.2)}C${f(w * 0.53)} ${f(h * 0.06)} ${f(w * 0.6)} 0 ${f(w * 0.7)} 0C${f(w * 0.85)} 0 ${f(w)} ${f(h * 0.1)} ${f(w)} ${f(h * 0.28)}C${f(w)} ${f(h * 0.45)} ${f(w * 0.9)} ${f(h * 0.7)} ${f(w / 2)} ${f(h)}Z` }),
  lightningBolt: (w, h) => ({ d: poly([[w * 0.39, 0], [w * 0.6, h * 0.28], [w * 0.51, h * 0.33], [w * 0.77, h * 0.52], [w * 0.69, h * 0.55], [w, h], [w * 0.4, h * 0.62], [w * 0.5, h * 0.58], [w * 0.2, h * 0.4], [w * 0.28, h * 0.36], [0, h * 0.1], [w * 0.25, h * 0.14]]) }),
  cloud: (w, h) => ({ d: `M${f(w * 0.2)} ${f(h * 0.75)}C${f(w * 0.02)} ${f(h * 0.78)} ${f(-w * 0.02)} ${f(h * 0.5)} ${f(w * 0.12)} ${f(h * 0.45)}C${f(w * 0.05)} ${f(h * 0.2)} ${f(w * 0.3)} ${f(h * 0.1)} ${f(w * 0.38)} ${f(h * 0.25)}C${f(w * 0.45)} ${f(h * 0.02)} ${f(w * 0.7)} ${f(h * 0.05)} ${f(w * 0.72)} ${f(h * 0.25)}C${f(w * 0.95)} ${f(h * 0.2)} ${f(w * 1.02)} ${f(h * 0.5)} ${f(w * 0.88)} ${f(h * 0.6)}C${f(w * 1.02)} ${f(h * 0.8)} ${f(w * 0.8)} ${f(h * 0.95)} ${f(w * 0.68)} ${f(h * 0.8)}C${f(w * 0.6)} ${f(h * 1.0)} ${f(w * 0.35)} ${f(h * 0.98)} ${f(w * 0.32)} ${f(h * 0.8)}C${f(w * 0.28)} ${f(h * 0.82)} ${f(w * 0.24)} ${f(h * 0.8)} ${f(w * 0.2)} ${f(h * 0.75)}Z` }),
  sun: (w, h) => { const pts = []; for (let i = 0; i < 16; i++) { const a = rad(-90 + 22.5 * i); const r = i % 2 ? 0.33 : 0.5; pts.push([w / 2 + w * 0.5 * (r / 0.5) * (i % 2 ? 0.66 : 1) * Math.cos(a) * (r > 0.4 ? 1 : 1), h / 2 + h * 0.5 * (i % 2 ? 0.66 : 1) * Math.sin(a)]); } return { d: poly(pts) + ellipse(w / 2, h / 2, w * 0.25, h * 0.25), evenodd: false }; },
  wedgeRectCallout: (w, h, A) => { const ax = w / 2 + (w * A('adj1', -20833)) / 100000; const ay = h / 2 + (h * A('adj2', 62500)) / 100000; const bottom = ay > h; const left = Math.min(w * 0.4, w * 0.2); const right = w * 0.4;
    return { d: bottom ? `M0 0H${f(w)}V${f(h)}H${f(w * 0.58)}L${f(ax)} ${f(ay)}L${f(w * 0.3)} ${f(h)}H0Z` : `M0 0H${f(w)}V${f(h)}H0Z`, text: [0, 0, w, h], approx: 'callout pointer is simplified' }; },
};

/** Preset geometry or null when the name is unknown. `adj` is { adj1: 50000, ... } read from <a:avLst>. */
export function presetGeometry(prst, w, h, adj) {
  const A = (n, d) => adjOf(adj, n, d);
  const p = PRESETS[prst]; const s = !p && SIMPLIFIED[prst];
  const r = p ? p(w, h, A) : s ? s(w, h, A) : null;
  if (!r) return null;
  const approx = r.approx || (s ? `preset shape "${prst}" is drawn with a simplified outline` : null);
  return { paths: [{ d: r.d, fill: !r.open, stroke: true, evenodd: !!r.evenodd }], textRect: r.text || null, approx, open: !!r.open };
}
export const hasPreset = (prst) => !!(PRESETS[prst] || SIMPLIFIED[prst]);
export const presetNames = () => [...Object.keys(PRESETS), ...Object.keys(SIMPLIFIED)];

// ---------- custom geometry (a:custGeom): guides, paths, arcs ----------
// Guide formulas work in the file's own units: lengths in EMU, angles in 60000ths of a degree. Guide values used as path coordinates are taken literally in the
// path's own coordinate space (<a:path w h>), as the specification and LibreOffice do; with no path size they are EMU.
const BUILTIN = (w, h) => { const ss = Math.min(w, h); const ls = Math.max(w, h); return { l: 0, t: 0, r: w, b: h, w, h, hc: w / 2, vc: h / 2, wd2: w / 2, hd2: h / 2, wd4: w / 4, hd4: h / 4, wd8: w / 8, hd8: h / 8, wd3: w / 3, hd3: h / 3, wd5: w / 5, hd5: h / 5, wd6: w / 6, hd6: h / 6, ss, ls, ssd2: ss / 2, ssd4: ss / 4, ssd6: ss / 6, ssd8: ss / 8, ssd16: ss / 16, ssd32: ss / 32, cd2: 10800000, cd4: 5400000, cd8: 2700000, '3cd4': 16200000, '3cd8': 8100000, '5cd8': 13500000, '7cd8': 18900000 }; };
function evalGuides(cg, wE, hE) {
  const env = BUILTIN(wE, hE);
  const v = (t) => { if (t == null) return 0; if (t in env) return env[t]; const n = Number(t); return Number.isFinite(n) ? n : 0; };
  for (const g of [...kids(kid(cg, 'avLst'), 'gd'), ...kids(kid(cg, 'gdLst'), 'gd')]) {
    const [op, ...a] = (g.getAttribute('fmla') || '').trim().split(/\s+/); const x = v(a[0]); const y = v(a[1]); const z = v(a[2]);
    let r = 0;
    switch (op) {
      case 'val': r = x; break; case '*/': r = z ? (x * y) / z : 0; break; case '+-': r = x + y - z; break; case '+/': r = z ? (x + y) / z : 0; break;
      case '?:': r = x > 0 ? y : z; break; case 'abs': r = Math.abs(x); break; case 'sqrt': r = Math.sqrt(Math.max(0, x)); break;
      case 'max': r = Math.max(x, y); break; case 'min': r = Math.min(x, y); break; case 'pin': r = pin(x, y, z); break; case 'mod': r = Math.sqrt(x * x + y * y + z * z); break;
      case 'sin': r = x * Math.sin(rad(y / 60000)); break; case 'cos': r = x * Math.cos(rad(y / 60000)); break; case 'tan': r = x * Math.tan(rad(y / 60000)); break;
      case 'at2': r = ((Math.atan2(y, x) * 180) / Math.PI) * 60000; break;
      case 'cat2': r = x * Math.cos(Math.atan2(z, y)); break; case 'sat2': r = x * Math.sin(Math.atan2(z, y)); break;
      default: r = 0;
    }
    env[g.getAttribute('name')] = r;
  }
  return env;
}
/** custGeom -> same result shape as presetGeometry. w, h are the shape size in points. */
export function customGeometry(cg, w, h) {
  const E = 12700;
  const env = evalGuides(cg, w * E, h * E);
  const raw = (t) => (t == null || t === '' ? 0 : t in env ? env[t] : Number(t) || 0);
  const paths = []; let approx = null;
  const rect = kid(cg, 'rect'); let textRect = null;
  if (rect) textRect = ['l', 't', 'r', 'b'].map((a) => raw(rect.getAttribute(a)) / E);
  for (const p of kids(kid(cg, 'pathLst'), 'path')) {
    const pw = Number(p.getAttribute('w')) || 0; const ph = Number(p.getAttribute('h')) || 0;
    const sx = pw ? w / pw : 1 / E; const sy = ph ? h / ph : 1 / E;
    const pt = (q) => [raw(q.getAttribute('x')) * sx, raw(q.getAttribute('y')) * sy];
    let d = ''; let cur = [0, 0]; let start = [0, 0];
    for (const c of p.children) {
      const n = c.localName;
      if (n === 'moveTo') { cur = start = pt(kid(c, 'pt')); d += `M${f(cur[0])} ${f(cur[1])}`; }
      else if (n === 'lnTo') { cur = pt(kid(c, 'pt')); d += `L${f(cur[0])} ${f(cur[1])}`; }
      else if (n === 'cubicBezTo') { const q = kids(c, 'pt').map(pt); if (q.length === 3) { d += `C${q.map((v) => `${f(v[0])} ${f(v[1])}`).join(' ')}`; cur = q[2]; } }
      else if (n === 'quadBezTo') { const q = kids(c, 'pt').map(pt); if (q.length === 2) { d += `Q${q.map((v) => `${f(v[0])} ${f(v[1])}`).join(' ')}`; cur = q[1]; } }
      else if (n === 'arcTo') {
        const a = arcTo(cur[0], cur[1], raw(c.getAttribute('wR')) * sx, raw(c.getAttribute('hR')) * sy, raw(c.getAttribute('stAng')) / 60000, raw(c.getAttribute('swAng')) / 60000);
        d += a.cmd; cur = [a.x, a.y];
      } else if (n === 'close') { d += 'Z'; cur = start; }
    }
    if (!d) continue;
    const fillMode = p.getAttribute('fill') || 'norm';
    if (!['norm', 'none'].includes(fillMode)) approx = `custom path fill mode "${fillMode}" (lighten / darken) is drawn as normal`;
    paths.push({ d, fill: fillMode !== 'none', stroke: p.getAttribute('stroke') !== '0' && p.getAttribute('stroke') !== 'false', evenodd: false });
  }
  if (!paths.length) return null;
  return { paths, textRect, approx, open: paths.every((p) => !p.fill) };
}
