// Node harness for the pure analysis layer (parser + design). Lets the analysis be tested without a browser.
// Needs: npm i jsdom (dev only). Not used by the app.
import { createRequire } from 'node:module';
import { readFile } from 'node:fs/promises';
import { fileURLToPath } from 'node:url';
import path from 'node:path';
const require = createRequire(import.meta.url);
const { JSDOM } = require('jsdom');
const root = path.resolve(path.dirname(fileURLToPath(import.meta.url)), '..');
const win = new JSDOM('').window;
globalThis.DOMParser = win.DOMParser;
globalThis.Node = win.Node;
// vendor/jszip.min.js is UMD; the package is "type":"module", so load it through a CommonJS-style shim.
{ const mod = { exports: {} }; new Function('module', 'exports', await readFile(path.join(root, 'vendor/jszip.min.js'), 'utf8'))(mod, mod.exports); globalThis.JSZip = mod.exports; }
const { parsePptx } = await import(path.join(root, 'js/pptx/parser.js'));
const design = await import(path.join(root, 'js/pptx/design.js'));
export const { extractDesign, DESIGN_VERSION, contrast } = design;
export async function analyse(file) {
  const buf = await readFile(file);
  const parsed = await parsePptx(new Uint8Array(buf));
  return { parsed, design: await extractDesign(parsed) };
}
