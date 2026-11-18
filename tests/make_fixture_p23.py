"""Fixture for the title / template-text / phantom-frame fixes (run from the repo root: python3 tests/make_fixture_p23.py).

  p23-template-text.pptx  16:9, 10 slides built like the real "FOOD" deck: plain text boxes (no placeholders), a brand name and a typed
                          page number at the same place on slides 2-10 (9 of 10), empty text boxes left inside groups, titles lower on the slide.
    1   cover: big title, subtitle, brand name at a DIFFERENT place (a cover variant, not template text)
    2   title in the upper half (baseline)
    3   80 pt title at 53% of the slide height (below the old 50% cut-off), body under it
    4   80 pt title at 68%, body above it
    5   26 pt heading at 60% over 24 pt body: lower but NOT dominant, so no title is inferred (and the brand name must not become the title)
    6   no large text at all: only 18 pt body; the brand name (21 pt, first text box) must not become the title
    7   empty text frames that ARE drawn (one with an outline, one with a fill) crossing the top edge: still decorative bleed;
        plus the number "42" at the same place on slides 7-9: same value every time, so not a page number
    8   as 7 (label only)
    9   as 7 (label only)
    10  closing
"""
from pptx import Presentation
from pptx.util import Emu, Pt
from pptx.dml.color import RGBColor

W, H = 12192000, 6858000
p = Presentation(); p.slide_width, p.slide_height = Emu(W), Emu(H)
BRAND = (0.874, 0.048, 0.095, 0.035)
NUM = (0.0, 0.912, 0.056, 0.069)


def box(shapes, fx, fy, fw, fh, t=None, pt=21, bold=False):
    b = shapes.add_textbox(Emu(int(W * fx)), Emu(int(H * fy)), Emu(int(W * fw)), Emu(int(H * fh)))
    b.text_frame.word_wrap = True
    if t is not None:
        b.text_frame.text = t
        for r in b.text_frame.paragraphs[0].runs: r.font.size = Pt(pt); r.font.bold = bold
    return b


def base(n, brand=True, number=True):
    s = p.slides.add_slide(p.slide_layouts[6])
    if brand: box(s.shapes, *BRAND, 'Acme & Co.', 21)                 # first text box on the slide, like the real deck
    if number: box(s.shapes, *NUM, f'{n:02d}', 42, True)
    g = s.shapes.add_group_shape()                                      # empty text boxes inside groups, one crossing the top edge
    box(g.shapes, 0.0, -0.02, 0.19, 0.03); box(g.shapes, 0.0, 0.0, 0.05, 0.02)
    return s


def para(k): return f'Sentence {k} about soil, water and harvest planning for the season ahead, written so that no two slides share a word count.' + ' More detail.' * k


s = p.slides.add_slide(p.slide_layouts[6])                              # 1
box(s.shapes, 0.29, 0.32, 0.42, 0.32, 'COVER', 150, True); box(s.shapes, 0.32, 0.61, 0.36, 0.1, 'A short subtitle for the cover', 28)
box(s.shapes, 0.464, 0.186, 0.095, 0.035, 'Acme & Co.', 21)

s = base(2); box(s.shapes, 0.056, 0.30, 0.8, 0.13, 'Upper title', 80, True); box(s.shapes, 0.056, 0.52, 0.8, 0.2, para(1))
s = base(3); box(s.shapes, 0.056, 0.53, 0.8, 0.13, 'Lower title at fifty three', 80, True); box(s.shapes, 0.056, 0.73, 0.8, 0.15, para(2))
s = base(4); box(s.shapes, 0.056, 0.68, 0.8, 0.13, 'Lowest title at sixty eight', 80, True); box(s.shapes, 0.056, 0.26, 0.8, 0.3, para(3))
s = base(5); box(s.shapes, 0.056, 0.60, 0.8, 0.08, 'Quiet lower heading', 26, True); box(s.shapes, 0.056, 0.20, 0.8, 0.3, para(4), 24)
s = base(6); box(s.shapes, 0.056, 0.30, 0.8, 0.3, para(5), 18)
for n in (7, 8, 9):
    s = base(n); box(s.shapes, 0.056, 0.30, 0.8, 0.13, f'Title number {n}', 80, True); box(s.shapes, 0.056, 0.52, 0.6, 0.2, para(n))
    box(s.shapes, 0.80, 0.55, 0.08, 0.08, '42', 30)
    if n == 7:
        o = box(s.shapes, 0.30, -0.02, 0.2, 0.05); o.line.color.rgb = RGBColor(0x20, 0x20, 0x20); o.line.width = Pt(2)
        f = box(s.shapes, 0.55, -0.02, 0.2, 0.05); f.fill.solid(); f.fill.fore_color.rgb = RGBColor(0xC8, 0x87, 0x3A)
s = base(10, number=False); box(s.shapes, 0.056, 0.40, 0.8, 0.13, 'Thank you', 80, True)
p.save('tests/fixtures/p23-template-text.pptx'); print('wrote p23-template-text.pptx')
