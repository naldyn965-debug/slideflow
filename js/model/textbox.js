// New text boxes (Phase 3.2.7-A) and new shapes (Phase 3.2.7-B). A text box or shape made in the preview is a RECORD in the temporary edit-state (js/model/edit-state.js: id, frame in EMU, one { text, rtl } per paragraph). The record is the
// box's "file": every time the slide's scene is built, the record is turned into a DETACHED <p:sp> element of the slide's own document (created here, never attached to the tree) and walked like any
// other shape, so selection, move, resize, style, rendering and text resolution see an ordinary text shape. The PPTX XML is never touched: the parsed document, the zip and the original shapes stay as
// they were; nothing is saved or exported.
export const NS_P = 'http://schemas.openxmlformats.org/presentationml/2006/main';
export const NS_A = 'http://schemas.openxmlformats.org/drawingml/2006/main';

const RTL = /[\u0590-\u05ff\u0600-\u06ff\u0700-\u08ff\ufb1d-\ufdff\ufe70-\ufefc]/;
const LTR = /[A-Za-z\u00c0-\u024f\u0370-\u03ff\u0400-\u04ff\u3040-\u30ff\u4e00-\u9fff\uac00-\ud7af]/;
/** Direction of a paragraph by its first strong character (what dir="auto" does): true = right to left. A paragraph without letters is left to right. */
export function paraRtl(text) {
  for (const ch of String(text)) { if (RTL.test(ch)) return true; if (LTR.test(ch)) return false; }
  return false;
}
/** A box whose text has nothing but white space, zero-width characters and bidi marks in it is empty. */
export const isBlankText = (texts) => texts.every((t) => /^[\s\u200b-\u200f\u202a-\u202e\u2060\ufeff]*$/.test(t));

/** The frame of a new box in EMU: 40% of the slide wide, about one line high, centred, each further box on the slide a little lower and to the right so none hides another. */
export function defaultFrame(slideW, slideH, n) {
  const w = Math.round(slideW * 0.4); const h = Math.round(slideH * 0.12); const step = Math.round(Math.min(slideW, slideH) * 0.04) * (n % 6);
  return { x: Math.min(Math.round((slideW - w) / 2) + step, slideW - w), y: Math.min(Math.round(slideH * 0.4) + step, slideH - h), w, h };
}

/** The kinds of element the preview can make: a text box and the five shapes of Phase 3.2.7-B (a line and an arrow are connectors). */
export const SHAPE_KINDS = ['rect', 'roundRect', 'ellipse', 'line', 'arrow'];
export const ADDED_KINDS = ['textbox', ...SHAPE_KINDS];
export const KIND_LABEL = { textbox: 'TextBox', rect: 'Rectangle', roundRect: 'Rounded Rectangle', ellipse: 'Oval', line: 'Straight Connector', arrow: 'Straight Arrow Connector' };
export const isLineKind = (k) => k === 'line' || k === 'arrow';
/** The frame of a new shape in EMU: a closed shape is 25% of the slide wide and 20% high, centred; a line or an arrow is 30% of the slide wide and flat (no height), at mid height. Each further element on the slide is offset a little down and right. */
export function shapeFrame(kind, slideW, slideH, n) {
  if (!isLineKind(kind)) { const w = Math.round(slideW * 0.25); const h = Math.round(slideH * 0.2); const step = Math.round(Math.min(slideW, slideH) * 0.04) * (n % 6); return { x: Math.min(Math.round((slideW - w) / 2) + step, slideW - w), y: Math.min(Math.round((slideH - h) / 2) + step, slideH - h), w, h }; }
  const w = Math.round(slideW * 0.3); const step = Math.round(Math.min(slideW, slideH) * 0.04) * (n % 6);
  return { x: Math.min(Math.round((slideW - w) / 2) + step, slideW - w), y: Math.min(Math.round(slideH / 2) + step, slideH), w, h: 0 };
}

const mk = (doc, tag, attrs = {}, ns = NS_A) => { const e = doc.createElementNS(ns, tag); for (const [k, v] of Object.entries(attrs)) e.setAttribute(k, String(v)); return e; };
/** A detached <p:sp> for the record `rec` ({ xmlId, name, x, y, w, h, paras: [{ text, rtl }] }), created in `doc` (the slide's document) and attached nowhere. */
export function textBoxElement(doc, rec) {
  const sp = mk(doc, 'p:sp', {}, NS_P);
  const nv = mk(doc, 'p:nvSpPr', {}, NS_P); nv.append(mk(doc, 'p:cNvPr', { id: rec.xmlId, name: rec.name }, NS_P), mk(doc, 'p:cNvSpPr', { txBox: 1 }, NS_P), mk(doc, 'p:nvPr', {}, NS_P));
  const pr = mk(doc, 'p:spPr', {}, NS_P); const xf = mk(doc, 'a:xfrm'); xf.append(mk(doc, 'a:off', { x: rec.x, y: rec.y }), mk(doc, 'a:ext', { cx: rec.w, cy: rec.h }));
  const geom = mk(doc, 'a:prstGeom', { prst: 'rect' }); geom.append(mk(doc, 'a:avLst')); pr.append(xf, geom, mk(doc, 'a:noFill'));
  const body = mk(doc, 'p:txBody', {}, NS_P); const bp = mk(doc, 'a:bodyPr', { wrap: 'square', lIns: 91440, tIns: 45720, rIns: 91440, bIns: 45720, rtlCol: 0, anchor: 't' }); bp.append(mk(doc, 'a:normAutofit'));
  body.append(bp, mk(doc, 'a:lstStyle'));
  for (const p of rec.paras) {
    const para = mk(doc, 'a:p'); const lang = p.rtl ? 'ar-SA' : 'en-US';
    if (p.rtl) para.append(mk(doc, 'a:pPr', { algn: 'r', rtl: 1 }));
    if (p.text !== '') { const r = mk(doc, 'a:r'); const t = mk(doc, 'a:t'); t.textContent = p.text; r.append(mk(doc, 'a:rPr', { lang, dirty: 0 }), t); para.append(r); } else para.append(mk(doc, 'a:endParaRPr', { lang, dirty: 0 }));
    body.append(para);
  }
  sp.append(nv, pr, body);
  return sp;
}
const scheme = (doc, val, shade) => { const c = mk(doc, 'a:schemeClr', { val }); if (shade) c.append(mk(doc, 'a:shade', { val: shade })); const f = mk(doc, 'a:solidFill'); f.append(c); return f; };
/** A detached <p:sp> (rectangle, rounded rectangle, ellipse) or <p:cxnSp> (line, arrow) for a shape record, created in `doc` and attached nowhere. Colours follow the deck's theme: a closed shape is filled accent 1 with a darker outline of the same accent (1 pt), a line or an arrow is 2.25 pt in the text colour; an arrow has a triangle at its end. */
export function shapeElement(doc, rec) {
  const line = isLineKind(rec.kind); const sp = mk(doc, line ? 'p:cxnSp' : 'p:sp', {}, NS_P);
  const nv = mk(doc, line ? 'p:nvCxnSpPr' : 'p:nvSpPr', {}, NS_P); nv.append(mk(doc, 'p:cNvPr', { id: rec.xmlId, name: rec.name }, NS_P), mk(doc, line ? 'p:cNvCxnSpPr' : 'p:cNvSpPr', {}, NS_P), mk(doc, 'p:nvPr', {}, NS_P));
  const pr = mk(doc, 'p:spPr', {}, NS_P); const xf = mk(doc, 'a:xfrm'); xf.append(mk(doc, 'a:off', { x: rec.x, y: rec.y }), mk(doc, 'a:ext', { cx: rec.w, cy: rec.h }));
  const geom = mk(doc, 'a:prstGeom', { prst: { rect: 'rect', roundRect: 'roundRect', ellipse: 'ellipse', line: 'line', arrow: 'straightConnector1' }[rec.kind] }); geom.append(mk(doc, 'a:avLst'));
  pr.append(xf, geom);
  if (!line) pr.append(scheme(doc, 'accent1'));
  const ln = mk(doc, 'a:ln', { w: line ? 28575 : 12700 }); ln.append(line ? scheme(doc, 'tx1') : scheme(doc, 'accent1', 50000));
  if (rec.kind === 'arrow') ln.append(mk(doc, 'a:tailEnd', { type: 'triangle', w: 'med', len: 'med' }));
  pr.append(ln); sp.append(nv, pr);
  return sp;
}
/** The detached element of any record (text box or shape). */
export const addedElement = (doc, rec) => (rec.kind && rec.kind !== 'textbox' ? shapeElement(doc, rec) : textBoxElement(doc, rec));

/** A detached <p:spTree> holding the boxes of `recs`, for the scene walker (never attached to a document). */
export function textBoxTree(doc, recs) { const tree = mk(doc, 'p:spTree', {}, NS_P); for (const r of recs) tree.append(addedElement(doc, r)); return tree; }
