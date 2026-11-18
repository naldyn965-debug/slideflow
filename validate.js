// Gatekeeper for uploads: cheap checks first, then a full parse. Throws PptxError with a readable message.
import { LIMITS } from '../config.js';
import { fmtBytes } from '../lib/format.js';
import { parsePptx, PptxError } from './parser.js';

export async function inspectPptx(file) {
  if (!/\.pptx$/i.test(file.name)) throw new PptxError('Only .pptx files are supported.', 'extension');
  if (!file.size) throw new PptxError('This file is empty.', 'empty');
  if (file.size > LIMITS.maxTemplateBytes) throw new PptxError(`This file is ${fmtBytes(file.size)}; the limit is ${fmtBytes(LIMITS.maxTemplateBytes)}.`, 'too-large');

  const head = new Uint8Array(await file.slice(0, 4).arrayBuffer());
  if (head[0] === 0xd0 && head[1] === 0xcf) throw new PptxError('This looks like an older .ppt file or a password-protected presentation. Remove the password or save it as .pptx and try again.', 'legacy-or-encrypted');
  if (!(head[0] === 0x50 && head[1] === 0x4b)) throw new PptxError('This file is not a valid PPTX (it is not a ZIP archive).', 'not-zip');

  return parsePptx(file);
}
