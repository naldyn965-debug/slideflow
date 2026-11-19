// Browser-free tests for the analysis layer (parser + design). Run: npm run test:analysis   (needs `npm i`, uses jsdom)
// Covers the Phase 2 ground truth that does not need a browser, plus Phase 2.1 precision fixes and Phase 2.2 classification/geometry.
import path from 'node:path';
import { existsSync } from 'node:fs';
import { fileURLToPath } from 'node:url';
import { analyse, analyseEdited, DESIGN_VERSION, contrast, sizeScale, sizeFloors } from './harness.mjs';
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

// Parser: a shape's own rotation is exposed as `rot` (degrees) so the preview can draw it where PowerPoint does; unrotated shapes carry no `rot` key.
const PGp = (await analyse(path.join(FX, 'p22-geometry.pptx'))).parsed;
const rotated = PGp.slides.flatMap((s, i) => s.shapes.filter((x) => 'rot' in x).map((x) => `${i + 1}:${x.kind}:${x.rot}`));
ok(rotated.join() === '1:shape:90,2:text:270' && PGp.slides[1].shapes.find((x) => x.rot).text === 'SIDE LABEL', `parser rot: p22-geometry has a rule rotated 90 and a text label rotated 270, nothing else (${rotated})`);

// The "every slide is bleed" case: each slide carries an accent bar and a footer band off the edges.
const PB = await D('p22-bleed-everywhere.pptx');
ok(PB.spacing.bleedSlides === 5 && PB.spacing.offSlideSlides === 0 && PB.slides.every((s) => s.bleed?.length && !s.offSlide && !s.overflow), 'p22 bleed-everywhere: all slides have decorative bleed and none has overflow');
ok(PB.slides.every((s) => s.bleedDetail.length === 2 && s.bleedDetail.every((b) => b.repeats === 5 && b.mode === 'beyond')), 'p22 bleed-everywhere: each bleed shape is shown to repeat on all 5 slides (template element)');
ok(PB.slides[1].bleedDetail.length === 2, 'p22 bleed-everywhere: the rotated rule inside the slide adds no third bleed shape');
ok(!finding(PB, 'content-overflow') && PB.modes.enhance.findings.length === 0, 'p22 bleed-everywhere: no overflow finding, no readability findings');
ok(near(PB.spacing.marginsPct.left, 8) && near(PB.spacing.marginsPct.top, 6) && near(PB.spacing.marginsPct.bottom, 30), `p22 bleed-everywhere: margins come from the text, not the bleed shapes (${JSON.stringify(PB.spacing.marginsPct)})`);
ok(PB.limitations.some((l) => /Rotation of a single shape is applied/.test(l)), 'p22 geometry: limitations describe rotation handling');

// ---------- Title extraction, repeated template text, phantom frames ----------
// tests/make_fixture_p23.py builds a deck like the real FOOD deck (brand name + typed page number on slides 2-10, empty text boxes inside groups,
// titles at 53% and 68% of the slide height). Before: every content slide was titled with the brand name, template text counted as content
// (words, margins), slides 3-4 had no inferred title, and the empty text boxes were reported as decorative bleed on 9 slides.
const T = await D('p23-template-text.pptx');
const ts = (n) => T.slides[n - 1];
ok(ts(1).title === 'COVER' && ts(2).title === 'Upper title' && ts(10).title === 'Thank you', `p23 title: inferred title text is the slide title (${ts(1).title} / ${ts(2).title} / ${ts(10).title})`);
ok(ts(3).title === 'Lower title at fifty three' && ts(4).title === 'Lowest title at sixty eight', `p23 title: an 80 pt title at 53% and at 68% of the height is detected (${ts(3).title} / ${ts(4).title})`);
ok(ts(5).title === '' && ts(6).title === '', `p23 title: a lower, non-dominant heading is not promoted, and the brand name is never the fallback (${JSON.stringify([ts(5).title, ts(6).title])})`);
ok(!T.slides.some((s) => /Acme/.test(s.title)), 'p23 title: recurring brand text is never a slide title');
ok(ts(2).templateText.map((x) => `${x.kind}:${x.text}`).join() === 'brand:Acme & Co.,pageNumber:02' && ts(9).templateText[1].text === '09' && ts(10).templateText.map((x) => x.kind).join() === 'brand', 'p23 template: brand name and page number found on slides 2-9, brand only on the closing slide');
ok(!ts(1).templateText.length && ts(1).words === 10, 'p23 template: the cover\'s brand name sits at another place, so it is not template text and still counts');
ok([2, 3, 4, 5, 6].map((n) => ts(n).words).join() === '26,31,33,33,32', `p23 template: brand name (3 words) and page number (1) leave the word counts (${[2, 3, 4, 5, 6].map((n) => ts(n).words)})`);
ok(near(ts(2).bounds.left, 5.6) && near(ts(2).bounds.top, 30) && near(ts(2).bounds.right, 14.4) && near(ts(2).bounds.bottom, 28), `p23 template: content margins come from title and body, not from the brand name and page number (${JSON.stringify(ts(2).bounds)})`);
ok(near(T.spacing.marginsPct.left, 5.6) && near(T.spacing.marginsPct.top, 30) && near(T.spacing.marginsPct.right, 14.4), `p23 template: template-wide margins follow (${JSON.stringify(T.spacing.marginsPct)})`);
ok(ts(2).counts.text === 4 && ts(2).counts.shape === 2, `p23 template: template text stays in the slide structure (${JSON.stringify(ts(2).counts)})`);
ok(!ts(7).templateText.some((x) => x.text === '42') && ts(7).words === 40 && near(ts(7).bounds.right, 12), 'p23 template: a number that does not follow the slide order ("42" on three slides) is content, not a page number');
ok(T.recurring.some((r) => r.kind === 'text' && r.label === 'Acme & Co.' && r.count === 9), 'p23 template: the Phase 2 "recurring" list is unchanged');
ok(cls(T) === 'title,text,text,text,text,text,text,text,text,closing', `p23 classes: unchanged by the fixes (${cls(T)})`);
ok([1, 2, 3, 4, 5, 6, 8, 9, 10].every((n) => !ts(n).bleed && !ts(n).bleedDetail) && !T.slides.some((s) => s.offSlide), 'p23 phantom: empty text boxes (no text, fill or outline) crossing the top edge are neither bleed nor content');
ok(ts(7).bleed?.join() === 'top' && ts(7).bleedDetail.length === 2 && T.spacing.bleedSlides === 1, `p23 phantom: empty text frames that ARE drawn (outline, fill) stay decorative bleed (${ts(7).bleedDetail?.length}, ${T.spacing.bleedSlides} slide)`);

// ---------- Cover detection: a dominant title is a cover however many small label boxes it has ----------
// The real FOOD cover has a 194 pt title and four small boxes (subtitle, tagline, URL and a brand name that sits at another place than on slides 2-14, so it is
// not template text). Before: more than 3 further text boxes made the first slide "text". Here slide 1 of p23-template-text (COVER 150 pt, a 28 pt subtitle and a 21 pt
// brand name) is extended in memory, so these checks run without the 27 MB deck.
const P23 = path.join(FX, 'p23-template-text.pptx');
const sp = (id, x, y, sz, text) => `<p:sp><p:nvSpPr><p:cNvPr id="${id}" name="Extra ${id}"/><p:cNvSpPr txBox="1"/><p:nvPr/></p:nvSpPr><p:spPr><a:xfrm><a:off x="${x}" y="${y}"/><a:ext cx="4389120" cy="400000"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom><a:noFill/></p:spPr><p:txBody><a:bodyPr wrap="square"><a:spAutoFit/></a:bodyPr><a:lstStyle/><a:p><a:r><a:rPr sz="${sz}" b="0"/><a:t>${text}</a:t></a:r></a:p></p:txBody></p:sp>`;
const tagline = (sz = 2100) => sp(5, 3901440, 5000000, sz, 'Learn why every meal matters');
const url = sp(6, 3901440, 5600000, 2100, 'www.example.com');
const cover = async (...steps) => (await analyseEdited(P23, { 'ppt/slides/slide1.xml': (xml) => steps.reduce((x, f) => f(x), xml) })).design.slides[0];
const add = (...xs) => (xml) => xml.replace('</p:spTree>', `${xs.join('')}</p:spTree>`);
const swap = (a, b) => (xml) => { if (!xml.includes(a)) throw new Error(`edit target not found: ${a}`); return xml.replace(a, b); };
const k1 = await cover(add(tagline(), url));
ok(k1.class === 'title' && k1.confidence === 'medium' && k1.reasons.some((r) => r === 'title text (150 pt) dominates the 4 other text boxes (28 pt at most)'), `cover: a 150 pt title over 4 small text boxes (28 pt at most) is a title (${k1.class}/${k1.confidence}: ${k1.reasons})`);
const k2 = await cover(swap('sz="15000"', 'sz="2800"'), add(tagline(), url));
ok(k2.class === 'text', `cover: 4 text boxes and no dominant title (title the same size as the subtitle) stays text (${k2.class})`);
const k3 = await cover(add(tagline(11000), url));
ok(k3.class === 'text', `cover: a title only 1.36x the next text (150 pt against 110 pt) does not dominate it, so the slide stays text (${k3.class})`);
const k4 = await cover(add(tagline(), url, sp(7, 3901440, 6000000, 1800, 'Each of the following lines is part of a long introductory paragraph that does not belong on a cover because it runs on and on well beyond what a title slide ever carries, naming food groups, meals, habits, drinks, shopping, cooking, storing and sharing.')));
ok(k4.class === 'text' && k4.reasons.every((r) => !/dominates/.test(r)), `cover: a dominant title with more than 30 words around it is still a text slide (${k4.class})`);
ok((await cover()).class === 'title' && (await cover()).reasons.every((r) => !/dominates/.test(r)), 'cover: a cover with 3 text boxes or fewer is decided as before (no dominance reason)');

// The real deck (27 MB, not committed): FOOD.pptx / "Brown and Green Modern Food Presentation". Set FOOD_PPTX or put it at tests/fixtures/FOOD.pptx.
const FOOD_FILE = process.env.FOOD_PPTX || path.join(FX, 'FOOD.pptx');
let FD = null;
if (existsSync(FOOD_FILE)) {
  FD = (await analyse(FOOD_FILE)).design;
  const f = (n) => FD.slides[n - 1];
  const titles = ['FOOD', 'What is Food?', 'Food Groups', 'Fruits', 'Vegetables', 'Proteins', 'Grains', 'Dairy', 'Different Types of Meals', 'Food from Around the World', 'Healthy vs. Unhealthy Food', 'Importance of Drinking Water', 'Food Waste Problem', 'How to Eat Well', 'THANK YOU'];
  ok(FD.slides.length === 15 && titles.every((t, k) => f(k + 1).title === t), `FOOD title: all 15 slides carry their content title (${FD.slides.map((s) => s.title).join(' | ')})`);
  ok(f(10).title === 'Food from Around the World' && f(11).title === 'Healthy vs. Unhealthy Food', 'FOOD S10/S11: titles at 53% and 68% of the height are detected');
  ok(!FD.slides.some((s) => s.title === 'Liceria & Co.'), 'FOOD title: "Liceria & Co." is never a slide title (it was the title of slides 2-14)');
  ok(FD.slides.slice(1, 14).every((s) => s.templateText.length === 2 && s.templateText[0].kind === 'brand' && s.templateText[0].text === 'Liceria & Co.' && s.templateText[1].kind === 'pageNumber' && Number(s.templateText[1].text) === s.n) && !f(1).templateText.length && !f(15).templateText.length, 'FOOD template: brand name + page number found on slides 2-14, nothing on the cover and closing slide');
  ok([f(2).words, f(10).words, f(11).words].join() === '64,53,52', `FOOD template: word counts leave out brand name and page number (S2/S10/S11 were 68/57/56; now ${[f(2).words, f(10).words, f(11).words]})`);
  ok(near(f(2).bounds.left, 5.6) && near(f(2).bounds.top, 5.1) && near(f(2).bounds.right, 5.6) && near(f(2).bounds.bottom, 10), `FOOD template: S2 content margins (were 0/4.8/3/1.9 from the brand name and page number): ${JSON.stringify(f(2).bounds)}`);
  ok(near(f(10).bounds.left, 5.6) && near(f(10).bounds.right, 13.2) && near(f(11).bounds.left, 6.3) && near(f(11).bounds.bottom, 18.7), `FOOD template: S10/S11 margins ${JSON.stringify(f(10).bounds)} ${JSON.stringify(f(11).bounds)}`);
  ok(near(FD.spacing.marginsPct.left, 5.6) && near(FD.spacing.marginsPct.top, 5.1) && near(FD.spacing.marginsPct.right, 12.6) && near(FD.spacing.marginsPct.bottom, 14.5), `FOOD template: template-wide margins ${JSON.stringify(FD.spacing.marginsPct)}`);
  ok(FD.recurring.length === 2 && FD.recurring.every((r) => r.count === 13) && FD.recurring.some((r) => r.kind === 'text' && r.label === 'Liceria & Co.'), 'FOOD template: the Phase 2 "recurring" list is unchanged (image + brand text, 13 slides each)');
  ok(f(1).title === 'FOOD' && f(1).words === 20 && near(f(1).bounds.left, 29.5) && near(f(1).bounds.top, 18.6) && near(f(1).bounds.right, 29.5) && near(f(1).bounds.bottom, 17.7) && f(1).counts.image === 2 && f(1).counts.text === 5 && !f(1).bleed, 'FOOD S1: cover title, words, margins and shape counts are unchanged (its brand text sits at another place than on slides 2-14)');
  ok(f(2).counts.shape === 8 && f(2).counts.text === 4 && f(10).counts.image === 4, 'FOOD template: counts of shapes, text boxes and pictures are unchanged (template text is flagged, not removed)');
  ok([2, 4, 7, 8, 12].every((n) => f(n).bleedDetail === undefined || f(n).bleedDetail.every((b) => b.kind === 'picture')) && !f(2).bleed && FD.slides.slice(1, 14).every((s) => !s.bleedDetail || s.bleedDetail.every((b) => b.kind === 'picture')), 'FOOD phantom: the empty text boxes inside the groups are no longer reported as top bleed (only pictures remain)');
  ok(f(10).bleed.join() === 'left,top,right,bottom' && f(11).bleed.join() === 'left,top' && f(10).bleedDetail.filter((b) => b.kind === 'picture').length === 3 && f(10).bleedDetail.filter((b) => b.kind === 'picture').every((b) => b.mode === 'flush'), `FOOD S10/S11: flush picture bleed is unchanged (${f(10).bleed}, ${f(11).bleed})`);
  ok(FD.spacing.bleedSlides === 12 && FD.spacing.offSlideSlides === 0 && !FD.slides.some((s) => s.offSlide), `FOOD phantom: bleed slides 13 -> ${FD.spacing.bleedSlides}, nothing off-slide`);
  // S1 cover. Its brand text sits at (46.4%, 18.6%) on slides 1 and 15 but at (87.4%, 4.8%) on slides 2-14: the first spot is shared by 2 of 15 slides
  // (5 needed), so it is not template text and the cover has 4 further text boxes beside its 194 pt title. The cover rule now accepts them when the title dominates.
  ok(cls(FD) === 'title,text,text,text,text,text,text,text,text,text,text,text,text,text,closing', `FOOD classes: only S1 changed, text -> title (${cls(FD)})`);
  ok(f(1).class === 'title' && f(1).confidence === 'medium' && f(1).reasons.some((r) => r === 'title text (194 pt) dominates the 4 other text boxes (35 pt at most)'), `FOOD S1: classified as a title, with the reason (${f(1).class}/${f(1).confidence}: ${f(1).reasons})`);
  ok(FD.slides.slice(1).map((s) => `${s.class}/${s.confidence}`).join() === [...Array(13).fill('text/medium'), 'closing/high'].join(), `FOOD S2-S15: classes and confidences exactly as before the cover fix (${FD.slides.slice(1).map((s) => `${s.class}/${s.confidence}`)})`);
  ok(FD.slides.slice(1).every((s) => s.reasons.every((r) => !/dominates/.test(r))), 'FOOD S2-S15: the cover rule leaves no trace outside slide 1');
  ok(FD.classification.counts.title === 1 && FD.classification.counts.text === 13 && FD.classification.counts.closing === 1, `FOOD classification counts: 1 title, 13 text, 1 closing (${JSON.stringify(FD.classification.counts)})`);

  // The 12 bleed slides. All of them are pictures touching an edge ("flush"); none is past an edge, so none is overflow. 21 pictures on slides 3-14.
  // Slide 12's photo stops 5,347 EMU short of the right edge, slide 9's two pictures overshoot the bottom by 207 and 414 EMU: all inside the 0.5% rounding tolerance.
  const flush = ['', '', 'right right', 'bottom', 'left right', 'right left', 'bottom', 'left+right+top', 'right+bottom bottom left+bottom', 'left+top top right+bottom', 'left+top top', 'right', 'right left', 'left+bottom', ''];
  ok(FD.slides.map((s) => (s.bleedDetail || []).map((b) => b.edges.join('+')).join(' ')).join('|') === flush.join('|'), `FOOD bleed: the exact pictures and edges behind the 12 slides (${FD.slides.map((s) => (s.bleedDetail || []).map((b) => b.edges.join('+')).join(' ')).join('|')})`);
  ok(FD.slides.filter((s) => s.bleed).map((s) => s.n).join() === '3,4,5,6,7,8,9,10,11,12,13,14' && FD.slides.flatMap((s) => s.bleedDetail || []).length === 21, 'FOOD bleed: slides 3-14 only (not the cover, S2 or the closing slide), 21 entries');
  ok(FD.slides.flatMap((s) => s.bleedDetail || []).every((b) => b.kind === 'picture' && b.mode === 'flush' && b.byPct === 0), 'FOOD bleed: every entry is a picture touching an edge; no shape, nothing past an edge');
  // Parser rot on the real deck: the S1/S15 backgrounds are stored as portrait boxes rotated -90, and the S7 photo is a rotated child of a group. Unrotated, they were drawn as a narrow strip / overlapping the text.
  const FP = (await analyse(FOOD_FILE)).parsed;
  const frot = FP.slides.flatMap((s, i) => s.shapes.filter((x) => 'rot' in x).map((x) => `${i + 1}:${x.rot}`));
  ok(frot.join() === '1:-90,3:90,3:-90,3:-90,7:90,10:90,14:90,15:-90' && FP.slides.flatMap((s) => s.shapes).filter((x) => 'rot' in x).every((x) => x.kind === 'image'), `FOOD parser rot: 8 rotated pictures, on slides 1, 3, 7, 10, 14, 15 (${frot})`);
  ok(FP.slides.flatMap((s) => s.shapes).filter((x) => !('rot' in x)).length === 209, 'FOOD parser rot: the other 209 shapes carry no rot key');

  ok(FD.spacing.offSlideSlides === 0 && FD.slides.every((s) => s.offSlide === undefined && s.overflow === undefined) && !finding(FD, 'content-overflow'), 'FOOD overflow: none (the 12 bleed slides are not overflow, and no overflow finding is raised)');

  // Genuine overflow is still reported on the real deck. Three shapes are moved in memory: S2 body text 15.5% past the right edge, S3 body text entirely outside the slide,
  // and the S4 photo group made 4% taller so it runs past the bottom edge (decorative: reported as "beyond", not as overflow).
  const mv = (from, to) => (xml) => { if (xml.split(from).length !== 2) throw new Error(`edit target not found exactly once: ${from}`); return xml.replace(from, to); };
  const FM = (await analyseEdited(FOOD_FILE, {
    'ppt/slides/slide2.xml': mv('<a:off x="8941058" y="5317102"/><a:ext cx="8318242" cy="2585085"/>', '<a:off x="12801600" y="5317102"/><a:ext cx="8318242" cy="2585085"/>'),
    'ppt/slides/slide3.xml': mv('<a:off x="1028700" y="5703570"/><a:ext cx="7345542" cy="2956560"/>', '<a:off x="19000000" y="5703570"/><a:ext cx="7345542" cy="2956560"/>'),
    'ppt/slides/slide4.xml': mv('<a:ext cx="9270863" cy="10058607"/>', '<a:ext cx="9270863" cy="10470087"/>'),
  })).design;
  const m = (n) => FM.slides[n - 1];
  ok(m(2).offSlide === true && m(2).overflow.length === 1 && m(2).overflow[0].kind === 'text' && m(2).overflow[0].edges.join() === 'right' && near(m(2).overflow[0].byPct, 15.5, 0.5) && !m(2).overflow[0].outside && /^Lorem ipsum/.test(m(2).overflow[0].sample), `FOOD overflow: S2 text running 15.5% past the right edge is reported (${JSON.stringify(m(2).overflow)})`);
  ok(m(3).offSlide === true && m(3).overflow[0].kind === 'text' && m(3).overflow[0].outside === true, `FOOD overflow: S3 text entirely outside the slide is reported (${JSON.stringify(m(3).overflow)})`);
  ok(m(4).bleedDetail[0].kind === 'picture' && m(4).bleedDetail[0].mode === 'beyond' && m(4).bleedDetail[0].edges.join() === 'bottom' && near(m(4).bleedDetail[0].byPct, 4, 0.2) && !m(4).offSlide, `FOOD bleed: a photo made 4% taller runs past the bottom edge and is reported as \"beyond\" (${JSON.stringify(m(4).bleedDetail[0])})`);
  const fo2 = finding(FM, 'content-overflow');
  ok(FM.spacing.offSlideSlides === 2 && fo2?.severity === 'warn' && fo2.slides.join() === '2,3' && FM.spacing.bleedSlides === 12, `FOOD overflow: summary and the content-overflow finding name slides 2 and 3 only; bleed slides stay 12 (${FM.spacing.offSlideSlides} / ${fo2?.slides} / ${FM.spacing.bleedSlides})`);
  ok(FD.slides.every((s, k) => [1, 2, 3].includes(k) || JSON.stringify(s) === JSON.stringify(FM.slides[k])), 'FOOD overflow: the three edits change slides 2-4 only; every other slide analyses exactly as before');
} else console.log(`SKIPPED: FOOD.pptx checks (set FOOD_PPTX=/path/to/Brown_and_Green_Modern_Food_Presentation.pptx or copy it to tests/fixtures/FOOD.pptx)`);

// ---------- contract ----------
const nested = (v) => (Array.isArray(v) ? v.some((x) => Array.isArray(x) || nested(x)) : v && typeof v === 'object' ? Object.values(v).some(nested) : false);
const all = [A, B, C, Dg, E, F, dark, dm, dx, S, PC, PA, PG, PB, T, ...(FD ? [FD] : [])];
ok(all.every((d) => !nested(d) && d.version === DESIGN_VERSION && d.version === 4 && d.schema === 'slideflow.design' && JSON.stringify(d).length < 100000), 'JSON: Firestore-safe, versioned (DESIGN_VERSION 4), small');
const v1Keys = ['n', 'title', 'layout', 'layoutType', 'class', 'confidence', 'reasons', 'background', 'words', 'counts', 'imageCoverage', 'minPt', 'maxPt'];
ok(all.every((d) => d.slides.every((s) => v1Keys.every((k) => k in s || s[k] === undefined) && ['title', 'section', 'text', 'image', 'comparison', 'chart-table', 'closing', 'other'].includes(s.class))), 'JSON: Phase 2 slide fields kept (changes are additive)');
ok(all.every((d) => d.slides.every((s) => s.templateText === undefined || s.templateText.every((x) => ['brand', 'pageNumber'].includes(x.kind) && typeof x.text === 'string'))), 'JSON: templateText entries are additive and typed');
ok(['left', 'right', 'top', 'bottom'].every((k) => k in dark.spacing.marginsPct) && 'gapPct' in dark.spacing && 'alignment' in dark.spacing && 'modes' in dark && 'limitations' in dark, 'JSON: Phase 2 top-level shape kept');
ok(all.every((d) => d.slides.every((s) => (s.imageRoles === undefined || ['background', 'content', 'decorative'].every((k) => Number.isInteger(s.imageRoles[k]))) && (s.overflow === undefined || s.offSlide === true) && (s.bleedDetail === undefined || (s.bleed && s.bleedDetail.every((b) => ['picture', 'shape'].includes(b.kind) && ['flush', 'beyond'].includes(b.mode)))))), 'JSON: Phase 2.2 fields (imageRoles, bleedDetail, overflow) are additive and consistent with bleed / offSlide');
const strip = (d) => JSON.stringify({ ...d, analyzedAt: 0 });
ok(strip(await D('p21-dark-explicit.pptx')) === strip(dark), 'analysis is deterministic');

// ---------- Phase 2.5 Step 2: Enhance rule coverage (tests and fixtures only; no production change) ----------
// Audit item #1. Every rule in buildModes() gets a fixture that trips it and one where it must stay silent, Preserve -> Enhance differences must equal
// Enhance.changes exactly, and Enhance must never lower a size or a margin. Fixtures: tests/make_fixture_p25.py.
// Floors at the reference 13.333 x 7.5 in slide (scale 1; since DESIGN_VERSION 4 every point floor is multiplied by the size scale, see the Step 3 section):
// body 16 pt; title max(28 pt, 1.25 x enhanced body); subtitle >= enhanced body; footer 12 pt; margins 5/5/5/4 % (never scaled); text colour 4.5:1 against the
// dominant background (replacement: the better of the theme's dk1 / lt1).
const FOODL = await D('p25-food-like.pptx');
const ES = await D('p25-enh-sizes.pptx'); const ER = await D('p25-enh-title-ratio.pptx'); const EM = await D('p25-enh-margins.pptx');
const ECL = await D('p25-enh-contrast-light.pptx'); const ECD = await D('p25-enh-contrast-dark.pptx'); const EF = await D('p25-enh-at-floors.pptx');
const changeKeys = (d) => d.modes.enhance.changes.map((c) => c.token).sort().join();
const chg = (d, token) => d.modes.enhance.changes.find((c) => c.token === token);
const leaves = (o, pre = '') => Object.entries(o).flatMap(([k, v]) => (v && typeof v === 'object' && !Array.isArray(v) ? leaves(v, `${pre}${k}.`) : [[`${pre}${k}`, v]]));
const byToken = (a, b) => a.token.localeCompare(b.token);
const diffOf = (d) => {
  const a = Object.fromEntries(leaves(d.modes.preserve.tokens)); const b = Object.fromEntries(leaves(d.modes.enhance.tokens));
  return [...new Set([...Object.keys(a), ...Object.keys(b)])].filter((k) => JSON.stringify(a[k]) !== JSON.stringify(b[k])).map((k) => ({ token: k, from: a[k], to: b[k] })).sort(byToken);
};
// Invariant: the token differences between Preserve and Enhance are exactly the listed changes (same token, same from, same to), each listed once.
const diffMatchesChanges = (d) => {
  const diff = diffOf(d); const ch = [...d.modes.enhance.changes].sort(byToken);
  const keys = (t) => leaves(t).map(([k]) => k).sort().join();
  return keys(d.modes.preserve.tokens) === keys(d.modes.enhance.tokens) && new Set(ch.map((c) => c.token)).size === ch.length && diff.length === ch.length
    && diff.every((x, i) => x.token === ch[i].token && JSON.stringify(x.from) === JSON.stringify(ch[i].from) && JSON.stringify(x.to) === JSON.stringify(ch[i].to))
    && ch.every((c) => JSON.stringify(c.from) !== JSON.stringify(c.to) && typeof c.reason === 'string' && c.reason.length > 0);
};
// Invariant: Enhance never lowers a size or a margin, never turns a missing value into one and never drops one.
const neverLowers = (d) => ['sizesPt', 'marginsPct'].every((g) => Object.keys(d.modes.preserve.tokens[g]).every((k) => {
  const p = d.modes.preserve.tokens[g][k]; const e = d.modes.enhance.tokens[g][k];
  return p == null ? e == null : typeof e === 'number' && e >= p;
})) && d.modes.enhance.changes.filter((c) => /^(sizesPt|marginsPct)\./.test(c.token)).every((c) => c.to > c.from);

// FOOD-like deck, 20 x 11.25 in (scale 1.5). Step 2 pinned the v3 behaviour (absolute floors: a 21 pt body on a 20 in wide slide was left alone) and said the
// expectation was meant to change once slide-size normalisation shipped. It has (DESIGN_VERSION 4): the body floor is 16 x 1.5 = 24 pt, so 21 -> 24 is the ONE change.
const fp = FOODL.modes.preserve.tokens;
ok(FOODL.slide.width === 18288000 && FOODL.slide.height === 10287000 && FOODL.slide.aspect === 1.78, `p25 FOOD-like: slide is 20 x 11.25 in (${FOODL.slide.width} x ${FOODL.slide.height} EMU)`);
ok(cls(FOODL) === 'title,text,text,text,text,text,text,closing', `p25 FOOD-like: cover, 6 text slides, closing (${cls(FOODL)})`);
ok(fp.sizesPt.title === 80 && fp.sizesPt.body === 21 && fp.sizesPt.subtitle === null && fp.sizesPt.footer === null && fp.fonts.heading === 'Inter Bold' && fp.fonts.body === 'Inter' && fp.colors.text === 'FFFFFF' && fp.colors.background === '211D19',
  `p25 FOOD-like: Preserve tokens match the real deck (80/21 pt, Inter Bold / Inter, white on #211D19): ${JSON.stringify(fp.sizesPt)}`);
const fm = FOODL.spacing.marginsPct;
ok(near(fm.left, 5.6) && near(fm.right, 12.6) && near(fm.top, 5.1) && near(fm.bottom, 14.5), `p25 FOOD-like: template margins match the real deck 5.6/12.6/5.1/14.5 (${fm.left}/${fm.right}/${fm.top}/${fm.bottom})`);
const bodyUp = (d) => d.modes.enhance.changes.length === 1 && d.modes.enhance.changes[0].token === 'sizesPt.body' && d.modes.enhance.changes[0].from === 21 && d.modes.enhance.changes[0].to === 24 && /at least 24 pt/.test(d.modes.enhance.changes[0].reason)
  && diffOf(d).length === 1 && diffOf(d)[0].token === 'sizesPt.body' && diffOf(d)[0].from === 21 && diffOf(d)[0].to === 24;
ok(bodyUp(FOODL) && FOODL.modes.enhance.tokens.sizesPt.title === 80 && JSON.stringify(FOODL.modes.enhance.tokens.marginsPct) === JSON.stringify(fm),
  `p25 FOOD-like: the only Enhance change is body 21 -> 24 pt (the intended v4 change; title 80 pt and every margin already meet their floors) (${changeKeys(FOODL)})`);
if (FD) {
  ok(FD.slide.width === 18288000 && FD.slide.height === 10287000 && bodyUp(FD) && JSON.stringify(FD.modes.preserve.tokens.sizesPt) === JSON.stringify(fp.sizesPt) && JSON.stringify(FD.modes.preserve.tokens.marginsPct) === JSON.stringify(fm)
    && JSON.stringify(FD.modes.enhance.tokens) === JSON.stringify(FOODL.modes.enhance.tokens),
    'FOOD (real deck): the only Enhance change is body 21 -> 24 pt, its Preserve tokens equal the FOOD-like fixture (80/21 pt, same margins) and so do its Enhance tokens');
} else console.log('SKIPPED: real-FOOD Enhance baseline (needs FOOD_PPTX or tests/fixtures/FOOD.pptx)');

// One positive-path fixture per rule. The isolation check (exact list of changed tokens) proves each fixture trips only the rule it targets.
ok(changeKeys(ES) === 'sizesPt.body,sizesPt.footer,sizesPt.subtitle,sizesPt.title', `p25 sizes: exactly the four size rules fire, no margin or colour change (${changeKeys(ES)})`);
ok(chg(ES, 'sizesPt.body')?.from === 12 && chg(ES, 'sizesPt.body').to === 16 && /at least 16 pt/.test(chg(ES, 'sizesPt.body').reason), 'p25 body rule: 12 pt body is raised to 16 pt');
ok(chg(ES, 'sizesPt.title')?.from === 20 && chg(ES, 'sizesPt.title').to === 28 && /28 pt/.test(chg(ES, 'sizesPt.title').reason), 'p25 title rule (28 pt floor): 20 pt title is raised to 28 pt');
ok(ES.modes.preserve.tokens.sizesPt.subtitle >= ES.modes.preserve.tokens.sizesPt.body && chg(ES, 'sizesPt.subtitle')?.from === 14 && chg(ES, 'sizesPt.subtitle').to === 16 && chg(ES, 'sizesPt.subtitle').to === ES.modes.enhance.tokens.sizesPt.body,
  'p25 subtitle rule: a 14 pt subtitle is not under the 12 pt body it has, yet is raised to 16 pt because it is compared with the ENHANCED body');
ok(chg(ES, 'sizesPt.footer')?.from === 9 && chg(ES, 'sizesPt.footer').to === 12 && /12 pt/.test(chg(ES, 'sizesPt.footer').reason), 'p25 footer rule: 9 pt footer placeholder is raised to 12 pt');
ok(changeKeys(ER) === 'sizesPt.title' && chg(ER, 'sizesPt.title')?.from === 30 && chg(ER, 'sizesPt.title').to === 35 && /1\.25x/.test(chg(ER, 'sizesPt.title').reason), `p25 title rule (1.25 x body): 30 pt title over 28 pt body becomes 35 pt and nothing else changes (${changeKeys(ER)})`);
const em = ['left', 'top', 'right', 'bottom'].map((k) => chg(EM, `marginsPct.${k}`));
ok(changeKeys(EM) === 'marginsPct.bottom,marginsPct.left,marginsPct.right,marginsPct.top' && em.every(Boolean) && em.map((c) => `${c.from}>${c.to}`).join() === '2>5,2>5,2>5,1>4', `p25 margin rules: 2/2/2/1 % become 5/5/5/4 % on all four sides, nothing else changes (${em.map((c) => c && `${c.from}>${c.to}`)})`);
const cL = chg(ECL, 'colors.text'); const cD = chg(ECD, 'colors.text');
const ratio = (c, bg) => (c ? `${contrast(c.to, bg).toFixed(1)}:1` : 'no change');   // safe in a failure message when the change is missing
ok(changeKeys(ECL) === 'colors.text' && cL?.from === 'BBBBBB' && cL.to === ECL.theme.colors.dk1 && contrast('BBBBBB', 'FFFFFF') < 4.5 && contrast(cL.to, 'FFFFFF') >= 4.5, `p25 colour rule (light): #BBBBBB on white (${contrast('BBBBBB', 'FFFFFF').toFixed(2)}:1) becomes the theme dk1 #${cL?.to} (${ratio(cL, 'FFFFFF')})`);
ok(changeKeys(ECD) === 'colors.text' && cD?.from === '444444' && cD.to === ECD.theme.colors.lt1 && contrast('444444', '211D19') < 4.5 && contrast(cD.to, '211D19') >= 4.5, `p25 colour rule (dark): #444444 on #211D19 (${contrast('444444', '211D19').toFixed(2)}:1) becomes the theme lt1 #${cD?.to} (${ratio(cD, '211D19')})`);
// Boundary: every value exactly ON its floor is left alone (floors are minimums).
const ef = EF.modes.preserve.tokens;
ok(ef.sizesPt.body === 16 && ef.sizesPt.title === 28 && ef.sizesPt.subtitle === 16 && ef.sizesPt.footer === 12 && ef.marginsPct.left === 5 && ef.marginsPct.right === 5 && ef.marginsPct.top === 5 && ef.marginsPct.bottom === 4, `p25 at-floors: the fixture really sits on every floor (${JSON.stringify(ef.sizesPt)} ${JSON.stringify(ef.marginsPct)})`);
ok(EF.modes.enhance.changes.length === 0 && diffOf(EF).length === 0, 'p25 at-floors: nothing changes when every value is exactly at its floor');

// Invariants over every design this suite analyses plus extra fixtures that trip rules.
const INV = [...all, FOODL, ES, ER, EM, ECL, ECD, EF, await D('wide.pptx'), await D('standard43.pptx'), await D('p21-style-layout-map.pptx'), await D('p22-geometry.pptx'), await D('p22-artwork.pptx')];
const failing = (pred) => INV.map((d, i) => [d, i]).filter(([d]) => !pred(d)).map(([, i]) => i);
ok(failing(diffMatchesChanges).length === 0, `INVARIANT: Preserve -> Enhance token differences are exactly Enhance.changes (token, from, to; each once) on ${INV.length} designs (failing: ${failing(diffMatchesChanges)})`);
ok(failing(neverLowers).length === 0, `INVARIANT: Enhance never lowers a size or margin, never adds or drops one, and every size/margin change goes up, on ${INV.length} designs (failing: ${failing(neverLowers)})`);
ok(INV.filter((d) => d.modes.enhance.changes.some((c) => /^sizesPt\./.test(c.token))).length >= 3, 'INVARIANT is not vacuous: at least 3 designs have a size raised');
const fired = [...new Set(INV.flatMap((d) => d.modes.enhance.changes.map((c) => c.token)))].sort().join();
ok(fired === 'colors.text,marginsPct.bottom,marginsPct.left,marginsPct.right,marginsPct.top,sizesPt.body,sizesPt.footer,sizesPt.subtitle,sizesPt.title',
  `COVERAGE tripwire: every Enhance token rule fires on at least one fixture and none else exists (${fired}). A new rule must come with a fixture and be added here`);
// The invariants are able to fail: a tampered copy must be rejected.
const tamper = (d, f) => { const c = JSON.parse(JSON.stringify(d)); f(c); return c; };
ok(!neverLowers(tamper(ES, (c) => { c.modes.enhance.tokens.sizesPt.body = 10; })) && !neverLowers(tamper(ES, (c) => { c.modes.enhance.tokens.sizesPt.footer = null; })) && !neverLowers(tamper(EM, (c) => { c.modes.enhance.tokens.marginsPct.left = 1; })), 'INVARIANT self-check: neverLowers rejects a lowered size, a dropped size and a lowered margin');
ok(!diffMatchesChanges(tamper(ES, (c) => { c.modes.enhance.changes.pop(); })) && !diffMatchesChanges(tamper(ES, (c) => { c.modes.enhance.tokens.sizesPt.title = 99; })) && !diffMatchesChanges(tamper(ES, (c) => { if (c.modes.enhance.changes[0]) c.modes.enhance.changes[0].to = 17; })), 'INVARIANT self-check: diffMatchesChanges rejects a missing change, an unlisted token difference and a wrong value');
ok([FOODL, ES, ER, EM, ECL, ECD, EF].every((d) => !nested(d) && d.version === DESIGN_VERSION && d.version === 4 && d.schema === 'slideflow.design' && JSON.stringify(d).length < 100000), 'p25 fixtures: Firestore-safe design JSON, DESIGN_VERSION is 4');
// ---------- Phase 2.5 Step 3: size-adaptive floors (DESIGN_VERSION 4) ----------
// Every POINT floor is multiplied by s = clamp(min(max(wr, hr), 1.2 * sqrt(wr * hr)), 0.25, 4), wr = width / 12192000, hr = height / 6858000, and rounded to
// 0.1 pt: body 16, title 28, footer 12, small-text 12, low-body-size 14, large text 18 pt (large bold 14 pt). NOT scaled: margins (a share of the slide), the
// title = 1.25 x body ratio, the 4.5:1 / 3:1 contrast ratios. Every expected value below is a hard-coded literal from an independent calculation (Python),
// never taken from the production function. Fixtures: tests/make_fixture_p25s3.py.
const nearTo = (a, b) => Math.abs(a - b) < 1e-6;
const FXP = (n) => path.join(FX, n);
const chs = (d) => d.modes.enhance.changes.map((c) => `${c.token}:${c.from}>${c.to}`).sort().join();
const lowAt = (d) => finding(d, 'low-contrast')?.slides.join();

// 1. The scale. [width, height, s]
const SCALES = [[12192000, 6858000, 1], [9144000, 6858000, 1], [9144000, 5143500, 0.75], [6096000, 3429000, 0.5], [18288000, 10287000, 1.5],
  [24384000, 6858000, 1.697056], [6858000, 9906000, 1.081665], [1219200, 685800, 0.25], [60000000, 40000000, 4]];
ok(SCALES.every(([w, h, e]) => nearTo(sizeScale(w, h), e)), `p25s3 scale: ${SCALES.map(([w, h]) => `${w}x${h}=${sizeScale(w, h).toFixed(6)}`).join(' ')}`);
ok(sizeScale(12192000, 6858000) === 1 && sizeScale(9144000, 6858000) === 1 && sizeScale(18288000, 10287000) === 1.5 && sizeScale(6096000, 3429000) === 0.5, 'p25s3 scale: 16:9 and 4:3 (13.33 / 10 in wide) are exactly 1, FOOD size exactly 1.5, half size exactly 0.5');
ok(sizeScale(24384000, 6858000) < 2 && sizeScale(6858000, 9906000) < 9906000 / 6858000, 'p25s3 scale: when wr and hr differ the 1.2 x sqrt(wr x hr) term decides (32:9 banner 1.70 < 2, portrait 1.08 < 1.44)');
ok([0.3, 0.5, 0.75, 1, 1.25, 1.5, 2, 3.5].every((k) => nearTo(sizeScale(k * 12192000, k * 6858000), k)), 'p25s3 scale: for a uniformly scaled 16:9 slide (wr = hr) the scale is that factor (the second term never binds)');
ok(sizeScale(1219200, 685800) === 0.25 && sizeScale(60000000, 40000000) === 4 && sizeScale(0, 0) === 1 && sizeScale(NaN, 5) === 1 && sizeScale(undefined, undefined) === 1, 'p25s3 scale: clamped to 0.25 and 4; a missing or zero size falls back to 1');
{ let prev = 0; let mono = true; for (let w = 1000000; w <= 80000000; w = Math.round(w * 1.07)) { const v = sizeScale(w, 6858000); if (v < prev) mono = false; prev = v; } ok(mono, 'p25s3 scale: never decreases as the slide gets wider'); }

// 2. The floors in points. [width, height, body, title, footer, small-text, low-body-size, large, large-bold]
const FLOORS = { 'small 0.5': [6096000, 3429000, 8, 14, 6, 6, 7, 9, 7], '10 x 5.625 in 0.75': [9144000, 5143500, 12, 21, 9, 9, 10.5, 13.5, 10.5], 'standard 16:9 1': [12192000, 6858000, 16, 28, 12, 12, 14, 18, 14],
  'standard 4:3 1': [9144000, 6858000, 16, 28, 12, 12, 14, 18, 14], 'large 1.5': [18288000, 10287000, 24, 42, 18, 18, 21, 27, 21], 'mixed wide 1.697': [24384000, 6858000, 27.2, 47.5, 20.4, 20.4, 23.8, 30.5, 23.8],
  'mixed tall 1.082': [6858000, 9906000, 17.3, 30.3, 13, 13, 15.1, 19.5, 15.1], 'clamp low 0.25': [1219200, 685800, 4, 7, 3, 3, 3.5, 4.5, 3.5], 'clamp high 4': [60000000, 40000000, 64, 112, 48, 48, 56, 72, 56] };
for (const [name, [w, h, ...e]] of Object.entries(FLOORS)) {
  const f = sizeFloors(w, h); const got = [f.body, f.title, f.footer, f.small, f.lowBody, f.large, f.largeBold];
  ok(got.join() === e.join(), `p25s3 floors (${name}): body/title/footer/small/low-body/large/large-bold = ${e.join(' / ')} (${got.join(' / ')})`);
}

// 3. Fixtures: slide size, Enhance changes and findings per size class
const S3 = {};
for (const n of ['small', 'small-at-floors', 'small-contrast', 'standard', 'standard-contrast', 'large', 'large-at-floors', 'large-contrast', 'large-margins', 'mixed-wide', 'mixed-wide-at-floors', 'mixed-wide-contrast', 'mixed-tall']) S3[n] = await D(`p25-s3-${n}.pptx`);
const DIMS = { small: '6096000x3429000', standard: '12192000x6858000', large: '18288000x10287000', 'mixed-wide': '24384000x6858000', 'mixed-tall': '6858000x9906000' };
ok(Object.entries(DIMS).every(([n, v]) => [n, `${n}-at-floors`, `${n}-contrast`, `${n}-margins`].filter((k) => S3[k]).every((k) => `${S3[k].slide.width}x${S3[k].slide.height}` === v)), 'p25s3 fixtures have the intended slide sizes (6.67x3.75, 13.33x7.5, 20x11.25, 26.67x7.5, 7.5x10.83 in)');
const SIZES = {   // every size rule fires; expected changes, then [small-text slides or undefined, low-body-size present]
  small: ['sizesPt.body:6>8,sizesPt.footer:4>6,sizesPt.subtitle:7>8,sizesPt.title:10>14', undefined, true, '{"title":10,"subtitle":7,"body":6,"footer":4}', 'under 6 pt', '8 pt or more'],
  standard: ['sizesPt.body:11>16,sizesPt.footer:9>12,sizesPt.subtitle:14>16,sizesPt.title:20>28', '1,2,3', true, '{"title":20,"subtitle":14,"body":11,"footer":9}', 'under 12 pt', '16 pt or more'],
  large: ['sizesPt.body:16>24,sizesPt.footer:14>18,sizesPt.subtitle:22>24,sizesPt.title:36>42', '1,2,3', true, '{"title":36,"subtitle":22,"body":16,"footer":14}', 'under 18 pt', '24 pt or more'],
  'mixed-wide': ['sizesPt.body:20>27.2,sizesPt.footer:14>20.4,sizesPt.subtitle:22>27.2,sizesPt.title:36>47.5', '1,2,3', true, '{"title":36,"subtitle":22,"body":20,"footer":14}', 'under 20.4 pt', '27.2 pt or more'],
  'mixed-tall': ['sizesPt.body:14>17.3,sizesPt.footer:10>13,sizesPt.subtitle:15>17.3,sizesPt.title:24>30.3', undefined, true, '{"title":24,"subtitle":15,"body":14,"footer":10}', null, '17.3 pt or more'],
};
for (const [n, [changes, smallSlides, lowBody, preserved, smallMsg, lowMsg]] of Object.entries(SIZES)) {
  const d = S3[n];
  ok(JSON.stringify(d.modes.preserve.tokens.sizesPt) === preserved && chs(d) === changes, `p25s3 ${n}: Preserve ${preserved}; Enhance changes exactly ${changes} (${chs(d)})`);
  ok(finding(d, 'small-text')?.slides.join() === smallSlides && (smallSlides === undefined || finding(d, 'small-text')?.message.includes(smallMsg)), `p25s3 ${n}: small-text ${smallSlides ? `flags slides ${smallSlides} (${smallMsg})` : 'does not fire'} (${finding(d, 'small-text')?.slides.join() ?? 'none'})`);
  ok(!!finding(d, 'low-body-size') === lowBody && finding(d, 'low-body-size')?.message.includes(lowMsg), `p25s3 ${n}: low-body-size fires and names the scaled ${lowMsg} (${finding(d, 'low-body-size')?.message})`);
}
ok(/at least 24 pt/.test(chg(S3.large, 'sizesPt.body')?.reason) && /at least 42 pt and 1\.25x body/.test(chg(S3.large, 'sizesPt.title')?.reason) && /18 pt or more/.test(chg(S3.large, 'sizesPt.footer')?.reason) && /scaled x1\.5 for this slide size/.test(chg(S3.large, 'sizesPt.body')?.reason),
  'p25s3 large: change reasons state the scaled floors (24 / 42 / 18 pt) and the 1.5 scale');
ok(!/scaled/.test(chg(S3.standard, 'sizesPt.body')?.reason) && /at least 16 pt/.test(chg(S3.standard, 'sizesPt.body')?.reason) && /at least 28 pt/.test(chg(S3.standard, 'sizesPt.title')?.reason) && /12 pt/.test(chg(S3.standard, 'sizesPt.footer')?.reason), 'p25s3 standard: scale 1 keeps the v3 wording and numbers, with no scale note');
ok(chg(S3['mixed-wide'], 'sizesPt.title')?.to === 47.5 && 27.2 * 1.25 < 47.5, 'p25s3 mixed-wide: the 47.5 pt title floor (28 x 1.697, 0.1 pt precision) decides, not 1.25 x the 27.2 pt body');

// 4. A value exactly ON a scaled floor is left alone; 0.1 pt under it fires (small, large and the fractional mixed-wide floors)
const AT = { small: ['{"title":14,"subtitle":8,"body":8,"footer":6}'], large: ['{"title":42,"subtitle":24,"body":24,"footer":18}'], 'mixed-wide': ['{"title":47.5,"subtitle":27.2,"body":27.2,"footer":20.4}'] };
for (const [n, [sizes]] of Object.entries(AT)) {
  const d = S3[`${n}-at-floors`]; const m = d.modes.preserve.tokens.marginsPct;
  ok(JSON.stringify(d.modes.preserve.tokens.sizesPt) === sizes && m.left === 5 && m.right === 5 && m.top === 5 && m.bottom === 4, `p25s3 ${n} at floors: the fixture sits exactly on every scaled floor ${sizes} and the 5/5/5/4 margins`);
  ok(d.modes.enhance.changes.length === 0 && d.modes.enhance.findings.length === 0 && diffOf(d).length === 0, `p25s3 ${n} at floors: nothing changes and no finding fires (a floor is a minimum)`);
}
const under = await analyseEdited(FXP('p25-s3-mixed-wide-at-floors.pptx'), Object.fromEntries([1, 2, 3].map((i) => [`ppt/slides/slide${i}.xml`, (x) => x.replaceAll('sz="4750"', 'sz="4740"').replaceAll('sz="2720"', 'sz="2710"').replaceAll('sz="2040"', 'sz="2030"')])));
ok(chs(under.design) === 'sizesPt.body:27.1>27.2,sizesPt.footer:20.3>20.4,sizesPt.subtitle:27.1>27.2,sizesPt.title:47.4>47.5', `p25s3 mixed-wide: 0.1 pt under each scaled floor raises all four (${chs(under.design)})`);
const smallAt = async (sz) => (await analyseEdited(FXP('p25-s3-mixed-wide.pptx'), Object.fromEntries([1, 2, 3].map((i) => [`ppt/slides/slide${i}.xml`, (x) => x.replaceAll('sz="2000"', `sz="${sz}"`)])))).design;
ok(finding(await smallAt(2030), 'small-text')?.slides.join() === '1,2,3' && finding(await smallAt(2040), 'small-text') === undefined, 'p25s3 mixed-wide: small-text boundary is the scaled 20.4 pt (20.3 pt flagged, 20.4 pt not)');
const ratioEdit = await analyseEdited(FXP('p25-s3-large-at-floors.pptx'), Object.fromEntries([1, 2, 3].map((i) => [`ppt/slides/slide${i}.xml`, (x) => x.replaceAll('sz="2400"', 'sz="4000"')])));
ok(chs(ratioEdit.design) === 'sizesPt.title:42>50', `p25s3 large: the 1.25 x body title ratio is not scaled (40 pt body makes a 42 pt title 50 pt; the scaled 42 pt floor alone would not) (${chs(ratioEdit.design)})`);

// 5. Margins are a share of the slide: never scaled
const lm = S3['large-margins']; const lmm = ['left', 'top', 'right', 'bottom'].map((k) => chg(lm, `marginsPct.${k}`));
ok(chs(lm) === 'marginsPct.bottom:1>4,marginsPct.left:2>5,marginsPct.right:2>5,marginsPct.top:2>5' && lmm.every(Boolean) && lmm.map((c) => `${c.from}>${c.to}`).join() === '2>5,2>5,2>5,1>4', `p25s3 large (scale 1.5): margins still 2/2/2/1 % -> 5/5/5/4 %, not scaled, and no size changes (${chs(lm)})`);

// 6. Contrast: slide 2 and 4 of each deck sit just under the scaled large-text boundary; ratios stay 4.5:1 / 3:1
const CON = { small: ['ok,low,ok,low', '[8.5,6.5]', 10], standard: ['ok,low,ok,low', '[17,13]', 18], large: ['ok,low,ok,low', '[24,20]', 27], 'mixed-wide': ['ok,low,ok,low', '[30,23.5]', 30.5] };
for (const [n, [statuses, sizes, bound]] of Object.entries(CON)) {
  const d = S3[`${n}-contrast`]; const f = finding(d, 'low-contrast');
  ok(d.slides.map((x) => x.contrast.status).join() === statuses && lowAt(d) === '2,4' && JSON.stringify(f?.evidence.map((e) => e.sizePt)) === sizes && f?.evidence.every((e) => e.ratio === 4.48),
    `p25s3 contrast (${n}): flagged slides 2 and 4 (sizes ${sizes}) around the scaled large-text boundary ${bound} pt (${d.slides.map((x) => x.contrast.status)})`);
}
ok(contrast('777777', 'FFFFFF') < 4.5 && contrast('777777', 'FFFFFF') >= 3, 'p25s3 contrast: the grey really is between 3:1 and 4.5:1, so only the large-text classification can decide');
ok(S3['large-contrast'].slides[1].contrast.status === 'low' && finding(S3['large-contrast'], 'low-contrast')?.evidence[0].sizePt === 24 && 24 >= 18, 'p25s3 contrast (large): 24 pt is "large" on a standard slide but normal text on a 20 in slide, so it is judged at 4.5:1');
ok(S3['small-contrast'].slides[0].contrast.status === 'ok' && finding(S3['small-contrast'], 'low-contrast')?.slides.indexOf(1) === -1, 'p25s3 contrast (small): 10 pt is normal text on a standard slide but "large" on a 6.67 in slide, so it is judged at 3:1 and passes');
const recolor = async (name, hex) => (await analyseEdited(FXP(`p25-s3-${name}.pptx`), Object.fromEntries([1, 2, 3, 4].map((i) => [`ppt/slides/slide${i}.xml`, (x) => x.replaceAll('777777', hex)])))).design;
for (const n of ['small', 'large', 'mixed-wide']) {
  const pass45 = await recolor(`${n}-contrast`, '767676'); const mid = await recolor(`${n}-contrast`, '888888'); const bad = await recolor(`${n}-contrast`, 'AAAAAA');
  ok(contrast('767676', 'FFFFFF') >= 4.5 && lowAt(pass45) === undefined, `p25s3 contrast (${n}): 4.54:1 passes everywhere, the 4.5:1 limit is unchanged`);
  ok(contrast('888888', 'FFFFFF') >= 3 && contrast('888888', 'FFFFFF') < 4.5 && lowAt(mid) === '2,4', `p25s3 contrast (${n}): 3.55:1 passes as large text only (flags slides 2 and 4), so 3:1 and 4.5:1 are unchanged (${lowAt(mid)})`);
  ok(contrast('AAAAAA', 'FFFFFF') < 3 && lowAt(bad) === '1,2,3,4', `p25s3 contrast (${n}): 2.32:1 fails even as large text (all four slides), the 3:1 limit is unchanged (${lowAt(bad)})`);
}

// 7. Real-world consequence of the approved formula: PptxGenJS's 10 x 5.625 in default has scale 0.75, so its 10 pt footer (floor 9 pt) is no longer raised
ok(Dg.slide.width === 9144000 && Dg.slide.height === 5143500 && Dg.typography.roles.footer?.sizePt === 10 && !chg(Dg, 'sizesPt.footer') && Dg.modes.enhance.tokens.sizesPt.footer === 10, 'p25s3 pptxgenjs (10 x 5.625 in, scale 0.75): the 10 pt footer is above the 9 pt floor and stays (v3 raised it to 12)');

// 8. Invariants and compatibility over every new fixture
const NEW = Object.values(S3);
ok(NEW.every(diffMatchesChanges) && NEW.every(neverLowers), `INVARIANT (Step 3): Preserve -> Enhance differences equal Enhance.changes, and nothing is lowered, on all ${NEW.length} new fixtures`);
ok(NEW.every((d) => !nested(d) && d.version === DESIGN_VERSION && d.version === 4 && d.schema === 'slideflow.design' && JSON.stringify(d).length < 100000), 'p25s3 fixtures: Firestore-safe design JSON, DESIGN_VERSION is 4');
ok(NEW.every((d) => ['left', 'right', 'top', 'bottom'].every((k) => k in d.spacing.marginsPct) && 'slide' in d && 'modes' in d && 'limitations' in d && d.slides.every((x) => 'class' in x && 'contrast' in x)), 'p25s3 fixtures: design JSON keeps the Phase 2 top-level and per-slide shape (additive compatibility)');
ok(strip(await D('p25-s3-mixed-wide.pptx')) === strip(S3['mixed-wide']), 'p25s3: analysis of a scaled slide is deterministic');

console.log(`${pass} passed, ${fail} failed`);
process.exit(fail ? 1 : 0);
