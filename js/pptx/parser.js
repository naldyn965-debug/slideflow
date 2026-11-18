// PPTX parser: JSZip + DOMParser over ECMA-376 PresentationML. Read-only; never executes or renders file content.
// Extracts: slide order, slide size, per-slide text, shape boxes, image references, layout names.
import { LIMITS } from '../config.js';

const NS_R = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
const CT_PPTX = 'application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml';
const CT_NAMES = {
  'application/vnd.ms-powerpoint.presentation.macroEnabled.main+xml': 'macro-enabled presentations (.pptm)',
  'application/vnd.openxmlformats-officedocument.presentationml.template.main+xml': 'templates (.potx)',
  'application/vnd.openxmlformats-officedocument.presentationml.slideshow.main+xml': 'slide shows (.ppsx)',
};
const IMG_MIME = { png: 'image/png', jpg: 'image/jpeg', jpeg: 'image/jpeg', gif: 'image/gif', webp: 'image/webp', svg: 'image/svg+xml' };

export class PptxError extends Error {
  constructor(message, code = 'invalid') { super(message); this.name = 'PptxError'; this.code = code; }
}

// ---- XML helpers (localName based, so namespace prefixes don't matter) ----
export const kids = (el, name) => (el ? [...el.children].filter((c) => c.localName === name) : []);
export const kid = (el, name) => kids(el, name)[0] || null;
export const desc = (el, name) => [...el.getElementsByTagNameNS('*', name)];
export const num = (el, a) => Number(el.getAttribute(a)) || 0;

export async function readXml(zip, path) {
  const f = zip.file(path);
  if (!f) return null;
  if (f._data && f._data.uncompressedSize > LIMITS.maxXmlBytes) throw new PptxError(`${path} is unusually large.`, 'too-large');
  const text = await f.async('string');
  if (text.length > LIMITS.maxXmlBytes) throw new PptxError(`${path} is unusually large.`, 'too-large');
  const doc = new DOMParser().parseFromString(text, 'application/xml');
  if (doc.getElementsByTagName('parsererror').length) throw new PptxError(`${path} contains invalid XML.`, 'bad-xml');
  return doc;
}

function resolvePath(fromPart, target) {
  if (target.startsWith('/')) return target.slice(1);
  const parts = fromPart.split('/').slice(0, -1);
  for (const seg of target.split('/')) {
    if (seg === '..') parts.pop();
    else if (seg && seg !== '.') parts.push(seg);
  }
  return parts.join('/');
}

export async function readRels(zip, part) {
  const i = part.lastIndexOf('/') + 1;
  const doc = await readXml(zip, `${part.slice(0, i)}_rels/${part.slice(i)}.rels`);
  const rels = new Map();
  if (!doc) return rels;
  for (const r of desc(doc, 'Relationship')) {
    const external = r.getAttribute('TargetMode') === 'External';
    const target = r.getAttribute('Target') || '';
    rels.set(r.getAttribute('Id'), {
      kind: (r.getAttribute('Type') || '').split('/').pop(),
      target: external ? target : resolvePath(part, target),
      external,
    });
  }
  return rels;
}

// ---- geometry ----
function readXfrm(x) {
  const off = kid(x, 'off');
  const ext = kid(x, 'ext');
  if (!off || !ext) return null;
  const box = { x: num(off, 'x'), y: num(off, 'y'), w: num(ext, 'cx'), h: num(ext, 'cy') };
  const chOff = kid(x, 'chOff');
  const chExt = kid(x, 'chExt');
  if (chOff && chExt) box.child = { x: num(chOff, 'x'), y: num(chOff, 'y'), w: num(chExt, 'cx'), h: num(chExt, 'cy') };
  return box;
}
export function xfrmOf(el) {
  const holder = el.localName === 'graphicFrame' ? el : kid(el, 'spPr') || kid(el, 'grpSpPr');
  return readXfrm(kid(holder, 'xfrm'));
}
export function intoGroup(box, g) {
  const c = g.child;
  if (!c) return box;
  const sx = c.w ? g.w / c.w : 1;
  const sy = c.h ? g.h / c.h : 1;
  return { x: g.x + (box.x - c.x) * sx, y: g.y + (box.y - c.y) * sy, w: box.w * sx, h: box.h * sy };
}

export function placeholderOf(el) {
  const nv = kid(el, 'nvSpPr') || kid(el, 'nvPicPr') || kid(el, 'nvGraphicFramePr');
  const ph = kid(kid(nv, 'nvPr'), 'ph');
  return ph ? { type: ph.getAttribute('type') || 'body', idx: ph.getAttribute('idx') } : null;
}
export const normType = (t) => (t === 'ctrTitle' ? 'title' : t);

// Placeholders without their own xfrm inherit position from the layout, then the master.
function inheritedBox(ph, docs) {
  for (const doc of docs) {
    if (!doc) continue;
    for (const sp of desc(doc, 'sp')) {
      const p = placeholderOf(sp);
      if (!p) continue;
      const same = ph.idx && p.idx ? ph.idx === p.idx : normType(ph.type) === normType(p.type);
      const box = same && xfrmOf(sp);
      if (box) return box;
    }
  }
  return null;
}

// ---- text ----
export function paragraphsText(container) {
  const body = kid(container, 'txBody');
  if (!body) return '';
  return kids(body, 'p')
    .map((p) => [...p.children].map((n) => (n.localName === 'br' ? '\n' : n.localName === 'r' || n.localName === 'fld' ? (kid(n, 't') || {}).textContent || '' : '')).join(''))
    .join('\n')
    .trim();
}
export function tableText(frame) {
  return desc(frame, 'tr').map((tr) => kids(tr, 'tc').map(paragraphsText).join('\t')).join('\n').trim();
}

// ---- slide shapes ----
function walkTree(tree, toSlide, ctx, out) {
  for (const el of tree.children) {
    const n = el.localName;
    if (n === 'grpSp') {
      const g = xfrmOf(el);
      walkTree(el, g ? (b) => toSlide(intoGroup(b, g)) : toSlide, ctx, out);
      continue;
    }
    if (!['sp', 'pic', 'graphicFrame'].includes(n)) continue;
    const ph = placeholderOf(el);
    const own = xfrmOf(el);
    const box = own ? toSlide(own) : ph ? inheritedBox(ph, ctx.inheritDocs) : null;
    const nv = kid(el, 'nvSpPr') || kid(el, 'nvPicPr') || kid(el, 'nvGraphicFramePr');
    const shape = { kind: 'shape', name: (kid(nv, 'cNvPr') || { getAttribute: () => '' }).getAttribute('name') || '', placeholder: ph ? normType(ph.type) : null, box, text: '', image: null };

    const blip = desc(el, 'blip')[0];
    const rel = blip && ctx.rels.get(blip.getAttributeNS(NS_R, 'embed') || blip.getAttributeNS(NS_R, 'link'));
    if (rel) shape.image = { target: rel.target, external: rel.external };

    if (n === 'pic') shape.kind = 'image';
    else if (n === 'graphicFrame') {
      shape.kind = desc(el, 'tbl').length ? 'table' : desc(el, 'chart').length ? 'chart' : 'object';
      if (shape.kind === 'table') shape.text = tableText(el);
    } else {
      shape.text = paragraphsText(el);
      if (shape.text) shape.kind = 'text';
      else if (shape.image) shape.kind = 'image';
    }
    out.push(shape);
  }
}

async function parseSlide(zip, part, index, cache, warnings) {
  const doc = await readXml(zip, part);
  if (!doc) { warnings.push(`Slide ${index + 1}: file ${part} is missing.`); return null; }
  const rels = await readRels(zip, part);
  const layoutRel = [...rels.values()].find((r) => r.kind === 'slideLayout');
  let layoutDoc = null;
  let masterDoc = null;
  let layoutName = null;
  if (layoutRel) {
    layoutDoc = cache.get(layoutRel.target) || (await readXml(zip, layoutRel.target));
    cache.set(layoutRel.target, layoutDoc);
    layoutName = layoutDoc ? (desc(layoutDoc, 'cSld')[0] || { getAttribute: () => null }).getAttribute('name') : null;
    const masterRel = [...(await readRels(zip, layoutRel.target)).values()].find((r) => r.kind === 'slideMaster');
    if (masterRel) {
      masterDoc = cache.get(masterRel.target) || (await readXml(zip, masterRel.target));
      cache.set(masterRel.target, masterDoc);
    }
  }
  const tree = desc(doc, 'spTree')[0];
  const shapes = [];
  if (tree) walkTree(tree, (b) => b, { rels, inheritDocs: [layoutDoc, masterDoc] }, shapes);

  const titleShape = shapes.find((s) => s.placeholder === 'title' && s.text) || shapes.find((s) => s.kind === 'text');
  const images = shapes.filter((s) => s.image).map((s) => s.image);
  for (const im of images) {
    if (im.external) warnings.push(`Slide ${index + 1}: picture is linked from outside the file and won't be available.`);
    else if (!zip.file(im.target)) warnings.push(`Slide ${index + 1}: picture ${im.target} is missing from the file.`);
  }
  return {
    index,
    part,
    hidden: doc.documentElement.getAttribute('show') === '0',
    layout: layoutName,
    title: titleShape ? titleShape.text.split('\n')[0].slice(0, 160) : '',
    shapes,
    images,
  };
}

/** Parse a .pptx (Blob, File, ArrayBuffer or Uint8Array). Throws PptxError with a user-readable message. */
export async function parsePptx(input) {
  const JSZip = globalThis.JSZip;
  if (!JSZip) throw new PptxError('The ZIP library did not load. Reload the page and try again.', 'no-zip');
  let zip;
  try { zip = await JSZip.loadAsync(input); } catch { throw new PptxError('This file is not a valid PPTX (it could not be opened as a ZIP archive).', 'not-zip'); }

  const total = Object.values(zip.files).reduce((s, f) => s + ((f._data && f._data.uncompressedSize) || 0), 0);
  if (total > LIMITS.maxUncompressedBytes) throw new PptxError('This file expands to an unusually large size and was rejected.', 'too-large');

  const ct = await readXml(zip, '[Content_Types].xml');
  if (!ct) throw new PptxError('This is a ZIP file, but not an Office document.', 'not-ooxml');
  const main = desc(ct, 'Override').find((o) => o.getAttribute('PartName') === '/ppt/presentation.xml');
  if (!main) throw new PptxError('This is not a PowerPoint presentation.', 'not-pptx');
  const mainType = main.getAttribute('ContentType');
  if (mainType !== CT_PPTX) throw new PptxError(`SlideFlow supports standard .pptx files only, not ${CT_NAMES[mainType] || 'this presentation type'}. Save as .pptx and try again.`, 'unsupported-type');

  const pres = await readXml(zip, 'ppt/presentation.xml');
  if (!pres) throw new PptxError('The presentation part is missing from this file.', 'not-pptx');
  const sz = desc(pres, 'sldSz')[0];
  if (!sz || !num(sz, 'cx') || !num(sz, 'cy')) throw new PptxError('The slide size is missing from this file.', 'bad-size');

  const presRels = await readRels(zip, 'ppt/presentation.xml');
  const warnings = [];
  const slideParts = desc(pres, 'sldId').map((s) => (presRels.get(s.getAttributeNS(NS_R, 'id')) || {}).target).filter((t, i) => { if (!t) warnings.push(`Slide entry ${i + 1} has no matching file and was skipped.`); return t; });
  if (!slideParts.length) throw new PptxError('This presentation has no slides.', 'no-slides');
  if (slideParts.length > LIMITS.maxSlides) throw new PptxError(`This presentation has ${slideParts.length} slides; the limit is ${LIMITS.maxSlides}.`, 'too-many-slides');

  const cache = new Map();
  const slides = [];
  for (const [i, part] of slideParts.entries()) {
    const s = await parseSlide(zip, part, i, cache, warnings);
    if (s) slides.push(s);
  }
  const mediaCount = Object.keys(zip.files).filter((n) => n.startsWith('ppt/media/') && !zip.files[n].dir).length;
  const counts = { chart: 0, table: 0, object: 0 };
  slides.forEach((s) => s.shapes.forEach((sh) => { if (sh.kind in counts) counts[sh.kind]++; }));
  if (counts.chart || counts.object) warnings.push(`Contains ${counts.chart} chart(s) and ${counts.object} embedded object(s) (SmartArt, OLE, video). They are listed but not rendered in the preview.`);

  return {
    slideCount: slides.length,
    width: num(sz, 'cx'),
    height: num(sz, 'cy'),
    mediaCount,
    imageRefs: slides.reduce((n, s) => n + s.images.length, 0),
    slides,
    warnings,
    zip, // kept for lazy image loading in the preview; not serialised
  };
}

/** The small, JSON-safe summary stored in Firestore / IndexedDB. */
export function summarize(parsed) {
  return {
    slideCount: parsed.slideCount,
    width: parsed.width,
    height: parsed.height,
    mediaCount: parsed.mediaCount,
    imageRefs: parsed.imageRefs,
    slideTitles: parsed.slides.map((s) => s.title.slice(0, 120)),
    warnings: parsed.warnings.slice(0, 20).map((w) => w.slice(0, 300)),
  };
}

/** Object URL for a bitmap/SVG media part, or null when the browser can't show that format (EMF, TIFF...). Caller revokes. */
export async function mediaUrl(zip, target) {
  const mime = IMG_MIME[target.split('.').pop().toLowerCase()];
  const f = mime && zip.file(target);
  if (!f) return null;
  return URL.createObjectURL(new Blob([await f.async('uint8array')], { type: mime }));
}
