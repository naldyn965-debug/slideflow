"""Phase 2.5 / Step 3 fixtures: size-adaptive floors. Ground truth is asserted in tests/analysis.mjs (section "Phase 2.5 Step 3").
Run from the repo root:  python3 tests/make_fixture_p25s3.py   (needs python-pptx; also run by `npm run fixtures`).

The floors are written for a 12192000 x 6858000 EMU slide and multiplied by s = clamp(min(max(wr, hr), 1.2 * sqrt(wr * hr)), 0.25, 4),
wr = width / 12192000, hr = height / 6858000. This script carries its OWN copy of the formula (floors() below) only to size the "at floors" decks; the
expected values in the tests are hard-coded literals, so the tests do not depend on the production function or on this copy.

  size class   slide                       s          body  title  footer/small  low-body  large / large-bold
  small        6096000 x 3429000 (6.67x3.75 in)    0.5        8     14     6        7        9 / 7
  standard     12192000 x 6858000 (13.33x7.5 in)   1          16    28     12       14       18 / 14
  large        18288000 x 10287000 (20x11.25 in)   1.5        24    42     18       21       27 / 21      (the real FOOD deck's size)
  mixed wide   24384000 x 6858000 (26.67x7.5 in)   1.697056   27.2  47.5   20.4     23.8     30.5 / 23.8  (32:9: wr 2, hr 1, the 1.2*sqrt term decides)
  mixed tall   6858000 x 9906000 (7.5x10.83 in)    1.081665   17.3  30.3   13.0     15.1     19.5 / 15.1  (portrait: wr 0.5625, hr 1.444, the 1.2*sqrt term decides)

  p25-s3-<size>.pptx            every size rule fires (body, title, subtitle, footer) plus the low-body-size finding (and small-text where noted).
  p25-s3-<size>-at-floors.pptx  every value exactly ON its scaled floor: a floor is a minimum, so nothing may change (small, large, mixed-wide).
  p25-s3-large-margins.pptx     large slide, every size already on its scaled floor, content frame at 2 / 2 / 2 / 1 %: the margin floors stay 5 / 5 / 5 / 4 %
                                (margins are a share of the slide and are not scaled).
  p25-s3-<size>-contrast.pptx   four slides of #777777 on white (4.48:1, just under 4.5:1) at sizes straddling the scaled large-text boundary, so
                                the large-text (3:1) vs normal-text (4.5:1) classification decides which slides are flagged: slides 2 and 4.

All values are explicit runs and explicit geometry, so the expected tokens do not depend on python-pptx's default master.
"""
import math
from lxml import etree
from pptx import Presentation
from pptx.util import Emu
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN

OUT = 'tests/fixtures/'
NS = 'xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"'
LOREM = ('Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod tempor incididunt ut labore et dolore magna aliqua. '
         'Ut enim ad minim veniam, quis nostrud exercitation ullamco laboris nisi ut aliquip ex ea commodo consequat. '
         'Duis aute irure dolor in reprehenderit in voluptate velit esse cillum dolore eu fugiat nulla pariatur.')

SMALL, STANDARD, LARGE = (6096000, 3429000), (12192000, 6858000), (18288000, 10287000)
WIDE, TALL = (24384000, 6858000), (6858000, 9906000)


def r1(x): return math.floor(x * 10 + 0.5) / 10          # JS Math.round semantics (half up), as the analysis uses


def floors(W, H):
    wr, hr = W / 12192000, H / 6858000
    s = max(0.25, min(4.0, min(max(wr, hr), 1.2 * math.sqrt(wr * hr))))
    f = lambda pt: r1(pt * s)
    return dict(s=s, body=f(16), title=f(28), footer=f(12), small=f(12), low_body=f(14), large=f(18), large_bold=f(14))


def rgb(h): return RGBColor.from_string(h)


def style(tf, pt, color=None, bold=False, align=None):
    for p in tf.paragraphs:
        if align is not None: p.alignment = align
        for r in p.runs:
            r.font.bold = bold
            r._r.get_or_add_rPr().set('sz', str(int(round(pt * 100))))   # exact centipoints (Pt() truncates, e.g. 20.4 pt -> 20.39)
            if color: r.font.color.rgb = rgb(color)


def background(slide, hexv):
    slide.background.fill.solid(); slide.background.fill.fore_color.rgb = rgb(hexv)


def place(shape, W, H, x, y, w, h):
    shape.left, shape.top, shape.width, shape.height = Emu(round(W * x)), Emu(round(H * y)), Emu(round(W * w)), Emu(round(H * h))


def textbox(slide, W, H, x, y, w, h, text, pt, **kw):
    b = slide.shapes.add_textbox(Emu(round(W * x)), Emu(round(H * y)), Emu(round(W * w)), Emu(round(H * h)))
    b.text_frame.word_wrap = True; b.text_frame.text = text
    style(b.text_frame, pt, **kw)
    return b


def footer(slide, W, H, text, pt):
    """A real footer placeholder (type ftr) with explicit geometry and size, so the footer role is assigned."""
    x, y, w, h = round(W * 0.05), round(H * 0.965), round(W * 0.30), round(H * 0.03)
    xml = (f'<p:sp {NS}><p:nvSpPr><p:cNvPr id="90" name="Footer Placeholder 90"/><p:cNvSpPr><a:spLocks noGrp="1"/></p:cNvSpPr>'
           f'<p:nvPr><p:ph type="ftr" sz="quarter" idx="11"/></p:nvPr></p:nvSpPr><p:spPr><a:xfrm><a:off x="{x}" y="{y}"/><a:ext cx="{w}" cy="{h}"/></a:xfrm></p:spPr>'
           f'<p:txBody><a:bodyPr/><a:lstStyle/><a:p><a:r><a:rPr lang="en-US" sz="{int(round(pt * 100))}"/><a:t>{text}</a:t></a:r></a:p></p:txBody></p:sp>')
    slide.shapes._spTree.append(etree.fromstring(xml))


def sized_deck(name, size, title_pt, body_pt, sub_pt, footer_pt, margins=(8, 8, 8, 8), slides=3):
    """Same construction as the Step 2 `enh_deck`, at any slide size: `slides` identical slides on the Title Slide layout; title and subtitle are real
    placeholders, the body is a plain text box; margins = (left, top, right, bottom) in % of the slide."""
    W, H = size
    l, t, r, b = margins
    p = Presentation(); p.slide_width, p.slide_height = Emu(W), Emu(H)
    for i in range(slides):
        s = p.slides.add_slide(p.slide_layouts[0])
        s.shapes.title.text = f'Heading {i + 1}'
        place(s.shapes.title, W, H, l / 100, t / 100, 0.40, 0.10)
        style(s.shapes.title.text_frame, title_pt, color='111111', bold=True, align=PP_ALIGN.LEFT)
        sub = s.placeholders[1]
        sub.text = 'A short subtitle line'
        place(sub, W, H, l / 100, (t + 12) / 100, 0.40, 0.08)
        style(sub.text_frame, sub_pt, color='222222', align=PP_ALIGN.LEFT)
        bw, bh = 0.40, 0.30
        textbox(s, W, H, (100 - r) / 100 - bw, (100 - b) / 100 - bh, bw, bh, LOREM, body_pt, color='222222', align=PP_ALIGN.LEFT)
        footer(s, W, H, 'Footer text', footer_pt)
    p.save(OUT + name); print('wrote', name)


def contrast_deck(name, size, runs):
    """One slide per (pt, bold): #777777 (4.48:1, under 4.5:1) on white. Large text only needs 3:1, so which slides are flagged depends on the
    scaled large-text boundary alone."""
    W, H = size
    p = Presentation(); p.slide_width, p.slide_height = Emu(W), Emu(H)
    for pt, bold in runs:
        s = p.slides.add_slide(p.slide_layouts[6]); background(s, 'FFFFFF')
        textbox(s, W, H, 0.10, 0.30, 0.80, 0.30, 'Sample text', pt, color='777777', bold=bold, align=PP_ALIGN.LEFT)
    p.save(OUT + name); print('wrote', name)


def at_floors(name, size):
    f = floors(*size)
    sized_deck(name, size, title_pt=f['title'], body_pt=f['body'], sub_pt=f['body'], footer_pt=f['footer'], margins=(5, 5, 5, 4))


# small (s = 0.5): every value under its scaled floor (8 / 14 / 8 / 6); body 6 is under the 7 pt low-body-size line
sized_deck('p25-s3-small.pptx', SMALL, title_pt=10, body_pt=6, sub_pt=7, footer_pt=4)
at_floors('p25-s3-small-at-floors.pptx', SMALL)
contrast_deck('p25-s3-small-contrast.pptx', SMALL, [(10, False), (8.5, False), (7, True), (6.5, True)])   # boundary 9 pt / bold 7 pt: ok, flagged, ok, flagged

# standard (s = 1): the floors are the v3 constants. Body 11 is under the 12 pt small-text line, so small-text and low-body-size both fire
sized_deck('p25-s3-standard.pptx', STANDARD, title_pt=20, body_pt=11, sub_pt=14, footer_pt=9)
contrast_deck('p25-s3-standard-contrast.pptx', STANDARD, [(18, False), (17, False), (14, True), (13, True)])  # boundary 18 pt / bold 14 pt

# large (s = 1.5): body 16 pt is fine at standard size but under the 18 pt small-text line here
sized_deck('p25-s3-large.pptx', LARGE, title_pt=36, body_pt=16, sub_pt=22, footer_pt=14)
at_floors('p25-s3-large-at-floors.pptx', LARGE)
contrast_deck('p25-s3-large-contrast.pptx', LARGE, [(27, False), (24, False), (21, True), (20, True)])        # boundary 27 pt / bold 21 pt (24 pt would be "large" at standard size)

# mixed sizes: width and height scale differently, so the 1.2 * sqrt(wr * hr) term decides and the floors are fractional (0.1 pt rounding)
sized_deck('p25-s3-mixed-wide.pptx', WIDE, title_pt=36, body_pt=20, sub_pt=22, footer_pt=14)               # s = 1.697056; body 20 is just under the 20.4 small-text line
at_floors('p25-s3-mixed-wide-at-floors.pptx', WIDE)
contrast_deck('p25-s3-mixed-wide-contrast.pptx', WIDE, [(31, False), (30, False), (24, True), (23.5, True)])  # boundary 30.5 pt / bold 23.8 pt
sized_deck('p25-s3-mixed-tall.pptx', TALL, title_pt=24, body_pt=14, sub_pt=15, footer_pt=10)               # s = 1.081665

# margins are percentages of the slide: not scaled. Sizes sit on the large floors, so only the four margin rules fire
fl = floors(*LARGE)
sized_deck('p25-s3-large-margins.pptx', LARGE, title_pt=fl['title'], body_pt=fl['body'], sub_pt=fl['body'], footer_pt=fl['footer'], margins=(2, 2, 2, 1))
