// Independent producer for fixture D (pptxgenjs). Needs: npm i pptxgenjs. Run from repo root: node tests/make_fixture_pptxgenjs.mjs
import pptxgen from 'pptxgenjs';
const PNG = 'image/png;base64,iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mNkYPhfDwAChwGA60e6kgAAAABJRU5ErkJggg==';
const p = new pptxgen(); p.layout = 'LAYOUT_16x9';
p.defineSlideMaster({ title: 'BRAND', background: { color: 'F2F2F2' }, objects: [{ rect: { x: 0, y: 0, w: '100%', h: 0.35, fill: { color: '0B6E4F' } } }, { text: { text: 'Fieldnote', options: { x: 8.3, y: 5.2, w: 1.5, h: 0.3, fontFace: 'Courier New', fontSize: 10, color: '555555' } } }], slideNumber: { x: 0.3, y: 5.2, fontFace: 'Courier New', fontSize: 10 } });
let s = p.addSlide({ masterName: 'BRAND' }); s.addText('Field Report 2026', { x: 0.7, y: 1.6, w: 8.6, h: 1, fontFace: 'Courier New', fontSize: 40, bold: true, color: '0B6E4F' }); s.addText('Annual summary', { x: 0.7, y: 2.7, w: 8.6, h: 0.6, fontFace: 'Courier New', fontSize: 20, color: '333333' });
s = p.addSlide({ masterName: 'BRAND' }); s.addText('Soil health improved across all plots because of rotation and cover crops this season.', { x: 0.7, y: 0.9, w: 8.6, h: 3, fontFace: 'Courier New', fontSize: 22, color: '222222' });
s = p.addSlide({ masterName: 'BRAND' }); s.addText('Yield data', { x: 0.7, y: 0.5, w: 8.6, h: 0.7, fontFace: 'Courier New', fontSize: 30, bold: true, color: '0B6E4F' });
s.addChart(p.charts.BAR, [{ name: 'Yield', labels: ['A', 'B', 'C'], values: [3, 5, 4] }], { x: 0.7, y: 1.3, w: 8.6, h: 3.8 });
s = p.addSlide({ masterName: 'BRAND' }); s.addText('Plot table', { x: 0.7, y: 0.5, w: 8.6, h: 0.7, fontFace: 'Courier New', fontSize: 30, bold: true, color: '0B6E4F' });
s.addTable([[{ text: 'Plot' }, { text: 'Yield' }], [{ text: 'A' }, { text: '3' }], [{ text: 'B' }, { text: '5' }]], { x: 0.7, y: 1.4, w: 8.6, h: 2, fontFace: 'Courier New', fontSize: 16 });
s = p.addSlide({ masterName: 'BRAND' }); s.addText('Thank you', { x: 0.7, y: 1.8, w: 8.6, h: 1, fontFace: 'Courier New', fontSize: 40, bold: true, color: '0B6E4F' }); s.addText('Questions?', { x: 0.7, y: 2.9, w: 8.6, h: 0.6, fontFace: 'Courier New', fontSize: 20, color: '333333' });
await p.writeFile({ fileName: 'tests/fixtures/p2-pptxgenjs.pptx' }); console.log('pptxgenjs fixture ok');
