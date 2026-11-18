// Browser-free tests for the analysis layer (parser + design). Run: npm run test:analysis   (needs `npm i`, uses jsdom)
// Covers the Phase 2 ground truth that does not need a browser, plus Phase 2.1 precision fixes.
import path from 'node:path';
import { fileURLToPath } from 'node:url';
import { analyse, DESIGN_VERSION } from './harness.mjs';
const FX = path.join(path.dirname(fileURLToPath(import.meta.url)), 'fixtures');
let pass = 0; let fail = 0;
const ok = (c, m) => { if (c) pass++; else { fail++; console.error('FAIL', m); } };
const D = async (n) => (await analyse(path.join(FX, n))).design;
const cls = (d) => d.slides.map((s) => s.class).join(',');
const near = (a, b, t = 0.2) => a != null && Math.abs(a - b) <= t;
const finding = (d, code) => d.modes.enhance.findings.find((f) => f.code === code);

// ---------- Phase 2 regression (same ground truth as tests/e2e.mjs) ----------
const A = await D('p2-corporate.pptx'); const B = await D('p2-explicit.pptx'); const C = await D('p2-arabic.pptx');
const Dg = await D('p2-pptxgenjs.pptx'); const E = await D('p2-no-theme.pptx'); const F = await D('p2-image-only.pptx');
ok(cls(A) === 'title,section,text,text,text,text,image,comparison,chart-table,chart-table,closing', `A classes (${cls(A)})`);
ok(A.slides[0].confidence === 'high' && A.slides[1].confidence === 'high' && A.slides[6].confidence === 'medium', 'A confidence levels');
ok(finding(A, 'small-text')?.slides.join() === '5' && finding(A, 'low-contrast')?.slides.join() === '6', 'A: small text on 5 and low contrast on 6 only');
ok(A.slides[3].minPt === 32 && A.slides[3].maxPt === 44, 'A: sizes inherited through master');
const mg = B.spacing.marginsPct;
ok(near(mg.left, 8) && near(mg.right, 8) && near(mg.top, 10) && near(mg.bottom, 15), `B margins 8/8/10/15 (${mg.left}/${mg.right}/${mg.top}/${mg.bottom})`);
ok(cls(B) === 'title,text,text,closing' && B.modes.enhance.findings.length === 0 && B.modes.enhance.changes.length === 0, 'B classes, no findings');
ok(cls(C) === 'title,text,closing' && C.typography.fonts[0].name === 'Cairo', 'C Arabic classes and cs font');
ok(cls(Dg) === 'title,text,chart-table,chart-table,closing', `D classes (${cls(Dg)})`);
ok(E.theme.name === null && E.limitations.some((l) => /No theme part/.test(l)), 'E: missing theme reported');
ok(cls(F) === 'image', 'F: single picture-only slide stays "image"');

// ---------- Phase 2.1: dark template ----------
// 15 slides on #211D19. 1 and 15: full-slide pictures. 8: genuinely unreadable grey text. 7: dark text on a light card (fine).
// 9: picture bleeding off three edges. 12 and 13: footer band picture bleeding off three edges. 5: translucent white card.
const dark = await D('p21-dark-explicit.pptx');
const dm = await D('p21-dark-clrmap.pptx');             // dark mode via slide colour-map override, background from bg1
const dx = await D('p21-dark-clrmap-explicit-bg.pptx'); // explicit #211D19 background, text colour from tx1 via the override: the reported all-slides failure
for (const [name, d] of [['explicit', dark], ['clrmap', dm], ['clrmap+explicit bg', dx]]) {
  const low = finding(d, 'low-contrast');
  ok(low?.slides.join() === '8', `contrast (${name}): only slide 8 is flagged (${low?.slides.join()})`);
  ok(!d.modes.enhance.changes.some((c) => c.token === 'colors.text'), `contrast (${name}): no bogus text-colour change proposed`);
  ok(d.slides[6].contrast.status === 'ok', `contrast (${name}): dark text on a light card is not flagged`);
  ok(d.slides[4].contrast.status === 'ok', `contrast (${name}): white text over a translucent card is not flagged`);
  ok([0, 9, 10, 14].every((i) => d.slides[i].contrast.status === 'not-evaluated'), `contrast (${name}): text over full-slide pictures is reported as not evaluated`);
  ok(d.limitations.some((l) => /Contrast was not evaluated/.test(l)), `contrast (${name}): not-evaluated share disclosed`);
}
ok(dx.slides.every((s) => s.colorMap === 'slide'), 'colour map: slide-level override detected on every slide');
ok(dx.modes.preserve.tokens.colors.text === 'FFFFFF', `colour map: body text resolves to white through tx1 -> lt1 (${dx.modes.preserve.tokens.colors.text})`);
const ev = finding(dark, 'low-contrast').evidence[0];
ok(ev.slide === 8 && ev.text === '3B3733' && ev.background === '211D19' && ev.ratio < 2, `contrast evidence names the slide and colours (${JSON.stringify(ev)})`);

// geometry
const negs = dark.slides.filter((s) => s.bounds && Object.values(s.bounds).some((v) => v < 0)).map((s) => s.n);
ok(negs.length === 0, `margins: no negative values (${negs.join()})`);
ok(near(dark.spacing.marginsPct.left, 8) && near(dark.spacing.marginsPct.right, 8) && near(dark.spacing.marginsPct.top, 10), `margins: median reflects the text frame (${JSON.stringify(dark.spacing.marginsPct)})`);
ok(!finding(dark, 'inconsistent-margins'), 'margins: bleed pictures no longer cause "inconsistent margins"');
ok(dark.slides[8].bleed?.includes('right') && dark.slides[8].bleed.includes('top') && dark.slides[8].bleed.includes('bottom'), `bleed: slide 9 picture reported (${dark.slides[8].bleed})`);
ok([11, 12].every((i) => ['left', 'right', 'bottom'].every((e) => dark.slides[i].bleed?.includes(e))), 'bleed: footer band on slides 12 and 13 reported');
ok(dark.slides[1].bleed?.join() === 'left,top', 'bleed: decorative accent bar running off the top-left corner is recognised (slide 2)');
ok(dark.spacing.bleedSlides === 6 && !dark.slides.some((s) => s.offSlide), `bleed: 6 slides counted (bar on 2,3,4,12,13 + pictures on 9,12,13), no text off the slide (${dark.spacing.bleedSlides})`);
ok(dark.slides[0].bounds && dark.slides[0].bounds.left === 8, 'margins: full-slide picture ignored, text frame measured on slide 1');

// classification
ok(dark.slides[0].class === 'title' && dark.slides[14].class === 'closing', `classes: full-slide-picture cover is a title and last slide is a closing (${dark.slides[0].class}/${dark.slides[14].class})`);
ok(dark.slides[10].class === 'text', `classes: long text over a full-slide picture is text, not image (${dark.slides[10].class})`);
ok(dark.slides[3].class === 'section' && dark.slides[13].class === 'closing', 'classes: lone big text box is a section divider; "Contact us" is a closing');
ok(dark.slides[8].class === 'image', 'classes: a picture-led slide (45% pictures) is still "image"');
ok(dark.slides.filter((s) => s.class === 'image').length === 1, 'classes: only the genuinely picture-led slide is "image"');
ok(dark.slides[0].reasons.some((r) => /background/.test(r)) && dark.slides[14].reasons.some((r) => /full-slide picture/.test(r)), 'classes: reasons explain the picture is a background');
ok(dark.slides[0].background.picture === true && dark.slides[0].contentImageCoverage === 0 && dark.slides[0].imageCoverage >= 0.99, 'classes: picture coverage split into total and content');

// text colour from <p:style>, layout-level colour map
const S = await D('p21-style-layout-map.pptx');
ok(S.slides[0].contrast.status === 'ok', 'style: white text from the shape style on a dark panel is not flagged');
ok(S.slides[1].contrast.status === 'low' && S.slides[1].contrast.ratio < 1.5, 'style: the same white text on a light panel is flagged');
ok(S.slides[2].colorMap === 'layout' && S.slides[2].contrast.status === 'ok', 'colour map: layout-level override applied');

ok(S.slides[3].offSlide === true && S.slides[3].bounds.right === 0 && S.spacing.offSlideSlides === 1, 'geometry: a text box running off the slide is flagged (offSlide) and its margin is clamped to 0, never negative');
ok(S.slides.filter((s) => s.offSlide).length === 1 && !S.slides.some((s) => s.bounds && Object.values(s.bounds).some((v) => v < 0)), 'geometry: no negative margins anywhere in the style/layout fixture');

// ---------- contract ----------
const nested = (v) => (Array.isArray(v) ? v.some((x) => Array.isArray(x) || nested(x)) : v && typeof v === 'object' ? Object.values(v).some(nested) : false);
const all = [A, B, C, Dg, E, F, dark, dm, dx, S];
ok(all.every((d) => !nested(d) && d.version === DESIGN_VERSION && d.version === 2 && d.schema === 'slideflow.design' && JSON.stringify(d).length < 100000), 'JSON: Firestore-safe, versioned 2, small');
const v1Keys = ['n', 'title', 'layout', 'layoutType', 'class', 'confidence', 'reasons', 'background', 'words', 'counts', 'imageCoverage', 'minPt', 'maxPt'];
ok(all.every((d) => d.slides.every((s) => v1Keys.every((k) => k in s || s[k] === undefined) && ['title', 'section', 'text', 'image', 'comparison', 'chart-table', 'closing', 'other'].includes(s.class))), 'JSON: Phase 2 slide fields kept (changes are additive)');
ok(['left', 'right', 'top', 'bottom'].every((k) => k in dark.spacing.marginsPct) && 'gapPct' in dark.spacing && 'alignment' in dark.spacing && 'modes' in dark && 'limitations' in dark, 'JSON: Phase 2 top-level shape kept');
const strip = (d) => JSON.stringify({ ...d, analyzedAt: 0 });
ok(strip(await D('p21-dark-explicit.pptx')) === strip(dark), 'analysis is deterministic');
console.log(`${pass} passed, ${fail} failed`);
process.exit(fail ? 1 : 0);
