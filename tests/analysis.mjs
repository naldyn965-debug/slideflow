// Browser-free tests for the analysis layer (parser + design). Run: npm run test:analysis   (needs `npm i`, uses jsdom)
// Covers the Phase 2 ground truth that does not need a browser, plus Phase 2.1 precision fixes and Phase 2.2 classification/geometry.
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


// ---------- Phase 2.2: slide types (text density, picture role, structure) ----------
// tests/make_fixture_p22.py: p22-classes has one slide per rule. Before 2.2 slides 2-3 and 7 were "comparison" (two columns) and 8-9 "image" (picture coverage alone).
const PC = await D('p22-classes.pptx');
const pcs = (n) => PC.slides[n - 1];
ok(cls(PC) === 'title,text,text,comparison,comparison,comparison,text,text,text,image,image,image,text,text,closing', `p22 classes: full sequence (${cls(PC)})`);
ok(pcs(2).class === 'text' && pcs(3).class === 'text', 'p22 comparison: two columns alone (layout "Two Content" or two free text boxes) are text, not a comparison');
ok(pcs(3).reasons.some((r) => /columns alone do not make a comparison/.test(r)), 'p22 comparison: the reasons say why two columns were not enough');
ok(pcs(7).class === 'text', 'p22 comparison: "compared with" inside running prose is not comparison wording');
ok(pcs(4).class === 'comparison' && pcs(4).confidence === 'medium' && pcs(4).reasons.some((r) => /mirrored columns, each led by its own short heading/.test(r)), 'p22 comparison: mirrored columns with their own headings are a medium-confidence comparison without any wording');
ok(pcs(5).class === 'comparison' && pcs(5).confidence === 'high' && pcs(5).reasons.some((r) => /comparison layout/.test(r)), 'p22 comparison: the "Comparison" layout is high confidence');
ok(pcs(6).class === 'comparison' && pcs(6).confidence === 'high' && pcs(6).reasons.some((r) => /title/.test(r)), 'p22 comparison: "vs" in the title with two columns is high confidence');
ok(pcs(13).class === 'text', 'p22 comparison: stacked headed blocks (an agenda) are a list, not a comparison');
ok(pcs(8).class === 'text' && near(pcs(8).contentImageCoverage, 0.36, 0.02), `p22 image: a 36% picture beside a 70-word paragraph is an illustrated text slide (${pcs(8).class}, ${pcs(8).contentImageCoverage})`);
ok(pcs(9).class === 'text' && pcs(9).contentImageCoverage >= 0.3 && pcs(9).confidence === 'medium', 'p22 image: a 31% picture with 34 words stays text, with medium confidence because the picture competes');
ok(pcs(8).reasons.some((r) => /so the text leads/.test(r)), 'p22 image: the reasons mention the picture and why text leads');
ok(pcs(10).class === 'image' && pcs(10).reasons.some((r) => /10 words of supporting text/.test(r)), 'p22 image: the same picture with a short caption is an image slide');
ok(pcs(11).class === 'image' && pcs(11).confidence === 'high' && pcs(11).reasons.some((r) => /picture layout/.test(r)), 'p22 image: the "Picture with Caption" layout counts as layout evidence');
ok(pcs(12).class === 'image' && pcs(12).imageRoles.content === 3 && pcs(12).reasons.some((r) => /gallery/.test(r)), 'p22 image: three pictures with a light caption are a gallery');
ok(PC.slides.every((s) => s.confidence !== 'low' || s.class === 'other'), 'p22 classes: no low-confidence guesses on clear slides');

const PA = await D('p22-artwork.pptx');
ok(cls(PA) === 'title,text,text,text,image,text', `p22 artwork: repeated side-panel art is decoration, a unique picture is content (${cls(PA)})`);
ok([1, 2, 3].every((i) => PA.slides[i].imageRoles.decorative === 1 && PA.slides[i].imageRoles.content === 0 && PA.slides[i].contentImageCoverage === 0 && PA.slides[i].imageCoverage > 0.2), 'p22 artwork: repeated picture is reported as decorative, total coverage kept, content coverage 0');
ok(PA.slides[4].imageRoles.content === 1 && PA.slides[5].imageRoles.decorative === 6 && PA.slides[5].class === 'text', 'p22 artwork: six logo-sized pictures are decoration, not a gallery');

// ---------- Phase 2.2: geometry (decorative bleed vs content overflow) ----------
// p22-geometry, 12 slides. Before 2.2: slides 1-2 (rotated shapes inside the slide) were bleed / off-slide, 3-4 and 10 (table/chart past the edge) were "bleed",
// 6 (frame overshoots by less than its inset) and 9 (empty placeholder) were off-slide / bleed.
const PG = await D('p22-geometry.pptx');
const g = (n) => PG.slides[n - 1];
ok(!g(1).bleed && !g(1).offSlide && !g(1).bleedDetail, 'p22 geometry: a vertical rule stored as a rotated bar is inside the slide (no bleed, no overflow)');
ok(!g(2).offSlide && !g(2).overflow && g(2).bounds.right > 1 && g(2).bounds.right < 4, `p22 geometry: a rotated side label is measured where it is drawn (right margin ${g(2).bounds.right}, not 0)`);
ok(g(3).offSlide && g(3).overflow[0].kind === 'table' && g(3).overflow[0].edges.join() === 'right' && near(g(3).overflow[0].byPct, 15, 0.5) && !g(3).bleed, 'p22 geometry: a table past the right edge is content overflow, not bleed');
ok(g(4).offSlide && g(4).overflow[0].kind === 'chart' && g(4).overflow[0].edges.join() === 'bottom' && near(g(4).overflow[0].byPct, 20, 0.5) && !g(4).bleed, 'p22 geometry: a chart past the bottom edge is content overflow, not bleed');
ok(g(3).bounds.right === 0 && g(3).bounds.bottom >= 0, 'p22 geometry: the overflowing table is still part of the content frame (clipped, never negative)');
ok(g(5).bleed?.join() === 'right,top,bottom' && g(5).bleedDetail[0].kind === 'picture' && g(5).bleedDetail[0].mode === 'flush' && !g(5).offSlide, 'p22 geometry: a photo flush with three edges is decorative bleed (flush), not overflow');
ok(!g(6).offSlide && !g(6).overflow, 'p22 geometry: a text frame overshooting by less than its own inset clips nothing');
ok(g(7).offSlide && g(7).overflow[0].kind === 'text' && g(7).overflow[0].edges.join() === 'right' && near(g(7).overflow[0].byPct, 19.3, 0.5) && /^Growers share/.test(g(7).overflow[0].sample), 'p22 geometry: text really running off the slide is overflow with the amount and a sample');
ok(g(8).bleed?.join() === 'right,bottom' && g(8).bleedDetail[0].kind === 'shape' && g(8).bleedDetail[0].mode === 'beyond' && near(g(8).bleedDetail[0].byPct, 30, 0.5) && !g(8).offSlide, 'p22 geometry: a decorative circle bleeding off a corner is intentional bleed (beyond, 30%)');
ok(!g(9).bleed && !g(9).offSlide, 'p22 geometry: an empty placeholder stretched past the edge is neither decoration nor content (it is not drawn)');
ok(g(10).bleed?.join() === 'left,top' && g(10).offSlide && g(10).overflow[0].kind === 'table', 'p22 geometry: one slide can have decorative bleed and content overflow, reported separately');
ok(g(11).offSlide && g(11).overflow[0].outside === true, 'p22 geometry: text entirely outside the slide is overflow');
ok(!g(12).bleed && !g(12).offSlide, 'p22 geometry: a plain slide has neither');
ok(PG.spacing.bleedSlides === 3 && PG.spacing.offSlideSlides === 5, `p22 geometry: summary counts decorative bleed (${PG.spacing.bleedSlides}) and overflow (${PG.spacing.offSlideSlides}) apart`);
const fo = finding(PG, 'content-overflow');
ok(fo?.severity === 'warn' && fo.slides.join() === '3,4,7,10,11' && fo.evidence.every((e) => e.ratio === undefined && e.kind && e.slide), 'p22 geometry: a "content-overflow" finding lists the overflowing slides with evidence (and never the decorative ones)');
ok(!PG.slides.some((s) => s.bounds && Object.values(s.bounds).some((v) => v < 0)), 'p22 geometry: no negative margins');

// The "every slide is bleed" case: each slide carries an accent bar and a footer band off the edges.
const PB = await D('p22-bleed-everywhere.pptx');
ok(PB.spacing.bleedSlides === 5 && PB.spacing.offSlideSlides === 0 && PB.slides.every((s) => s.bleed?.length && !s.offSlide && !s.overflow), 'p22 bleed-everywhere: all slides have decorative bleed and none has overflow');
ok(PB.slides.every((s) => s.bleedDetail.length === 2 && s.bleedDetail.every((b) => b.repeats === 5 && b.mode === 'beyond')), 'p22 bleed-everywhere: each bleed shape is shown to repeat on all 5 slides (template element)');
ok(PB.slides[1].bleedDetail.length === 2, 'p22 bleed-everywhere: the rotated rule inside the slide adds no third bleed shape');
ok(!finding(PB, 'content-overflow') && PB.modes.enhance.findings.length === 0, 'p22 bleed-everywhere: no overflow finding, no readability findings');
ok(near(PB.spacing.marginsPct.left, 8) && near(PB.spacing.marginsPct.top, 6) && near(PB.spacing.marginsPct.bottom, 30), `p22 bleed-everywhere: margins come from the text, not the bleed shapes (${JSON.stringify(PB.spacing.marginsPct)})`);
ok(PB.limitations.some((l) => /Rotation of a single shape is applied/.test(l)), 'p22 geometry: limitations describe rotation handling');

// ---------- contract ----------
const nested = (v) => (Array.isArray(v) ? v.some((x) => Array.isArray(x) || nested(x)) : v && typeof v === 'object' ? Object.values(v).some(nested) : false);
const all = [A, B, C, Dg, E, F, dark, dm, dx, S, PC, PA, PG, PB];
ok(all.every((d) => !nested(d) && d.version === DESIGN_VERSION && d.version === 3 && d.schema === 'slideflow.design' && JSON.stringify(d).length < 100000), 'JSON: Firestore-safe, versioned 2, small');
const v1Keys = ['n', 'title', 'layout', 'layoutType', 'class', 'confidence', 'reasons', 'background', 'words', 'counts', 'imageCoverage', 'minPt', 'maxPt'];
ok(all.every((d) => d.slides.every((s) => v1Keys.every((k) => k in s || s[k] === undefined) && ['title', 'section', 'text', 'image', 'comparison', 'chart-table', 'closing', 'other'].includes(s.class))), 'JSON: Phase 2 slide fields kept (changes are additive)');
ok(['left', 'right', 'top', 'bottom'].every((k) => k in dark.spacing.marginsPct) && 'gapPct' in dark.spacing && 'alignment' in dark.spacing && 'modes' in dark && 'limitations' in dark, 'JSON: Phase 2 top-level shape kept');
ok(all.every((d) => d.slides.every((s) => (s.imageRoles === undefined || ['background', 'content', 'decorative'].every((k) => Number.isInteger(s.imageRoles[k]))) && (s.overflow === undefined || s.offSlide === true) && (s.bleedDetail === undefined || (s.bleed && s.bleedDetail.every((b) => ['picture', 'shape'].includes(b.kind) && ['flush', 'beyond'].includes(b.mode)))))), 'JSON: Phase 2.2 fields (imageRoles, bleedDetail, overflow) are additive and consistent with bleed / offSlide');
const strip = (d) => JSON.stringify({ ...d, analyzedAt: 0 });
ok(strip(await D('p21-dark-explicit.pptx')) === strip(dark), 'analysis is deterministic');
console.log(`${pass} passed, ${fail} failed`);
process.exit(fail ? 1 : 0);
