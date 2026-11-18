"""Phase 2.2 fixtures: slide-type classification (image vs text, comparison vs plain columns) and geometry
(decorative bleed vs content overflow). Ground truth is asserted in tests/analysis.mjs.
Run from the repo root:  python3 tests/make_fixture_p22.py   (needs python-pptx; also run by `npm run fixtures`).

  p22-classes.pptx          4:3, python-pptx default layouts + free text boxes: one slide per classification rule
  p22-artwork.pptx          a large picture repeated across slides is template artwork, a unique one is content
  p22-geometry.pptx         16:9: rotated shapes, tables/charts past the edge, flush photos, empty placeholders, text vs frame
  p22-bleed-everywhere.pptx 16:9: every slide has decorative bleed, none has overflow (the "all slides are bleed" case)
"""
import io, struct, zlib
from pptx import Presentation
from pptx.util import Emu, Pt, Inches
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE
from pptx.chart.data import CategoryChartData
from pptx.enum.chart import XL_CHART_TYPE

OUT = 'tests/fixtures/'


def png(rgb, w=40, h=24):
    raw = b''.join(b'\x00' + bytes(rgb) * w for _ in range(h))
    def chunk(t, d): return struct.pack('>I', len(d)) + t + d + struct.pack('>I', zlib.crc32(t + d) & 0xffffffff)
    return b'\x89PNG\r\n\x1a\n' + chunk(b'IHDR', struct.pack('>IIBBBBB', w, h, 8, 2, 0, 0, 0)) + chunk(b'IDAT', zlib.compress(raw)) + chunk(b'IEND', b'')


_n = [0]
def fresh_png():
    """A different picture every call (python-pptx stores identical bytes once, which would make every picture 'the same artwork')."""
    _n[0] += 1
    return png(((_n[0] * 37) % 256, (_n[0] * 91) % 256, (_n[0] * 53) % 256))


class Deck:
    def __init__(self, W, H):
        self.p = Presentation(); self.W, self.H = W, H
        self.p.slide_width, self.p.slide_height = Emu(W), Emu(H)

    def slide(self, layout=6):
        return self.p.slides.add_slide(self.p.slide_layouts[layout])

    def text(self, s, fx, fy, fw, fh, t, pt=20, bold=False, rot=None):
        b = s.shapes.add_textbox(Emu(int(self.W * fx)), Emu(int(self.H * fy)), Emu(int(self.W * fw)), Emu(int(self.H * fh)))
        b.text_frame.word_wrap = True; b.text_frame.text = t
        for r in b.text_frame.paragraphs[0].runs: r.font.size = Pt(pt); r.font.bold = bold
        if rot: b.rotation = rot
        return b

    def title(self, s, t): return self.text(s, 0.08, 0.06, 0.84, 0.12, t, 36, bold=True)

    def pic(self, s, fx, fy, fw, fh, blob=None):
        return s.shapes.add_picture(io.BytesIO(blob or fresh_png()), Emu(int(self.W * fx)), Emu(int(self.H * fy)), Emu(int(self.W * fw)), Emu(int(self.H * fh)))

    def rect(self, s, fx, fy, fw, fh, hexv='C8873A', rot=None, shape=MSO_SHAPE.RECTANGLE):
        r = s.shapes.add_shape(shape, Emu(int(self.W * fx)), Emu(int(self.H * fy)), Emu(int(self.W * fw)), Emu(int(self.H * fh)))
        r.fill.solid(); r.fill.fore_color.rgb = RGBColor.from_string(hexv); r.line.fill.background()
        if rot: r.rotation = rot
        return r

    def save(self, name): self.p.save(OUT + name); print('wrote', name)


def words(n, seed='alpha'):
    base = ('growers share soil data weekly so every crew can plan irrigation harvest and packing around real weather and field '
            'conditions instead of guesses made months earlier').split()
    return ' '.join((base * 10)[:n]).capitalize() + '.'


# =====================================================================================================
# p22-classes: one slide per rule. 4:3 so the default template's layout placeholders sit inside the slide.
# Expected classes (analysis.mjs): title, text, text, comparison(medium), comparison(high), comparison(high), text, text, text, image, image,
#                                  image(gallery), text, text, closing
# =====================================================================================================
d = Deck(9144000, 6858000)
s = d.slide(0); s.shapes.title.text = 'Field Guide 2026'; s.placeholders[1].text = 'Practical notes for growers'                           # 1 title
s = d.slide(3); s.shapes.title.text = 'Our approach'                                                                                       # 2 "Two Content" layout, two body placeholders
s.placeholders[1].text = words(14); s.placeholders[2].text = words(16)
s = d.slide(); d.title(s, 'Overview'); d.text(s, 0.08, 0.28, 0.40, 0.5, words(22)); d.text(s, 0.52, 0.28, 0.40, 0.5, words(24))               # 3 two free columns, nothing else
s = d.slide(); d.title(s, 'Plans')                                                                                                         # 4 mirrored columns, each with its own heading, no wording
for fx, head, body in [(0.08, 'Starter', words(13)), (0.52, 'Professional', words(15))]:
    d.text(s, fx, 0.26, 0.40, 0.08, head, 24, bold=True); d.text(s, fx, 0.36, 0.40, 0.4, body)
s = d.slide(4); s.shapes.title.text = 'Plans'                                                                                              # 5 PowerPoint "Comparison" layout
for idx, t in {1: 'Build', 2: words(13), 3: 'Buy', 4: words(12)}.items(): s.placeholders[idx].text = t
s = d.slide(); d.title(s, 'Remote vs office'); d.text(s, 0.08, 0.28, 0.40, 0.5, words(20)); d.text(s, 0.52, 0.28, 0.40, 0.5, words(22))   # 6 wording in the title
s = d.slide(); d.title(s, 'Results')                                                                                                       # 7 "compared" only in running prose
d.text(s, 0.08, 0.28, 0.40, 0.5, 'Yields were compared with last season across all plots and rose steadily ' + words(10)); d.text(s, 0.52, 0.28, 0.40, 0.5, words(24))
s = d.slide(); d.title(s, 'Our regions'); d.pic(s, 0.04, 0.22, 0.5, 0.72); d.text(s, 0.58, 0.26, 0.36, 0.6, words(70), 16)                    # 8 36% picture + a 70-word paragraph
s = d.slide(); d.title(s, 'Soil and water'); d.pic(s, 0.06, 0.24, 0.56, 0.55); d.text(s, 0.66, 0.30, 0.28, 0.4, words(34), 14)               # 9 31% picture + 34 words (not picture-led)
s = d.slide(); d.title(s, 'North field'); d.pic(s, 0.06, 0.24, 0.62, 0.62); d.text(s, 0.72, 0.3, 0.22, 0.3, words(10), 14)                   # 10 38% picture + short caption
s = d.slide(8); s.shapes.title.text = 'Barn'; s.placeholders[1].insert_picture(io.BytesIO(fresh_png())); s.placeholders[2].text = words(12)   # 11 "Picture with Caption" layout
s = d.slide(); d.title(s, 'Gallery')                                                                                                       # 12 three pictures + caption
for fx in (0.06, 0.37, 0.68): d.pic(s, fx, 0.26, 0.22, 0.40)
d.text(s, 0.08, 0.72, 0.84, 0.16, words(30), 14)
s = d.slide(); d.title(s, 'Agenda')                                                                                                        # 13 stacked headed blocks: structure is a list, not a comparison
for k, head in enumerate(['Welcome', 'Results', 'Plans']):
    d.text(s, 0.08, 0.24 + 0.2 * k, 0.5, 0.06, head, 22, bold=True); d.text(s, 0.08, 0.30 + 0.2 * k, 0.5, 0.1, words(12), 16)
s = d.slide(); d.title(s, 'Notes'); d.text(s, 0.08, 0.26, 0.84, 0.5, words(40))                                                            # 14 plain text
s = d.slide(0); s.shapes.title.text = 'Thank you'; s.placeholders[1].text = 'Questions?'                                                   # 15 closing
d.save('p22-classes.pptx')

# =====================================================================================================
# p22-artwork: picture roles. Slides 2-4 repeat the same side-panel artwork (template art, decorative); slide 5 has a unique picture (content).
# =====================================================================================================
d = Deck(9144000, 6858000)
art = png((120, 90, 60))
s = d.slide(0); s.shapes.title.text = 'Brand deck'; s.placeholders[1].text = 'Artwork test'
for t in ('Mission', 'Method', 'Reach'):
    s = d.slide(); d.title(s, t); d.pic(s, 0.62, 0.24, 0.34, 0.66, art); d.text(s, 0.08, 0.3, 0.45, 0.3, words(12), 16)                       # 34% x 66% = 22% of the slide, same place
s = d.slide(); d.title(s, 'Harvest'); d.pic(s, 0.06, 0.24, 0.62, 0.62); d.text(s, 0.72, 0.3, 0.22, 0.3, words(10), 14)
s = d.slide(); d.title(s, 'Small marks')                                                                                                    # six logo-sized pictures: decoration, not a gallery
for k in range(6): d.pic(s, 0.08 + 0.08 * k, 0.3, 0.05, 0.05)
d.text(s, 0.08, 0.5, 0.8, 0.2, words(30), 16)
d.save('p22-artwork.pptx')

# =====================================================================================================
# p22-geometry: 16:9. See analysis.mjs for the expected bleed / overflow per slide.
# =====================================================================================================
d = Deck(12192000, 6858000)
W, H = d.W, d.H
def head(s, t): d.title(s, t); d.text(s, 0.08, 0.30, 0.6, 0.2, words(14), 18)
s = d.slide(); head(s, 'Rotated rule')                                                  # 1 vertical rule built from a rotated horizontal bar, inside the slide
r = d.rect(s, 0, 0, 0.7 * H / W, 0.004 * W / H, rot=90); r.left = Emu(int(W * .03 - H * .35)); r.top = Emu(int(H * .5 - W * .002))
s = d.slide(); head(s, 'Rotated label')                                                 # 2 rotated side label, inside the slide
t = d.text(s, 0, 0, 0.5 * H / W, 0.06, 'SIDE LABEL', 12, rot=270); t.left = Emu(int(W * .96 - H * .25)); t.top = Emu(int(H * .5 - H * .03))
s = d.slide(); head(s, 'Table past the edge')                                            # 3 table running off the right edge: content overflow
s.shapes.add_table(3, 4, Emu(int(W * .30)), Emu(int(H * .55)), Emu(int(W * .85)), Emu(int(H * .25)))
s = d.slide(); head(s, 'Chart past the edge')                                            # 4 chart running off the bottom edge: content overflow
cd = CategoryChartData(); cd.categories = ['A', 'B', 'C']; cd.add_series('S', (1, 2, 3))
s.shapes.add_chart(XL_CHART_TYPE.COLUMN_CLUSTERED, Emu(int(W * .10)), Emu(int(H * .45)), Emu(int(W * .5)), Emu(int(H * .75)), cd)
s = d.slide(); d.title(s, 'Flush photo'); d.pic(s, 0.55, 0, 0.45, 1.0); d.text(s, 0.08, 0.30, 0.4, 0.3, words(14), 18)   # 5 photo flush with three edges: decorative bleed
s = d.slide(); d.title(s, 'Frame vs text')                                                # 6 frame overshoots by less than its own inset: the text is still inside
d.text(s, 0.5, 0.3, 0.5075, 0.3, words(14), 18)
s = d.slide(); d.title(s, 'Text past the edge'); d.text(s, 0.7, 0.3, 0.5, 0.3, words(14), 18)   # 7 text box really running off the slide: content overflow
s = d.slide(); head(s, 'Glow'); d.rect(s, 0.72, 0.5, 0.45, 0.8, 'F3EDE4', shape=MSO_SHAPE.OVAL)  # 8 decorative circle bleeding off the corner
s = d.slide(1); s.shapes.title.text = 'Empty placeholder'                                # 9 layout "Title and Content" with its body left empty and stretched past the edge
body = s.placeholders[1]; body.left, body.top, body.width, body.height = Emu(int(W * .3)), Emu(int(H * .3)), Emu(int(W * .9)), Emu(int(H * .9))
s = d.slide(); head(s, 'Bar and table'); d.rect(s, -0.02, -0.02, 0.2, 0.05)             # 10 decorative bar AND a table past the edge: reported separately
s.shapes.add_table(2, 3, Emu(int(W * .25)), Emu(int(H * .55)), Emu(int(W * .9)), Emu(int(H * .2)))
s = d.slide(); head(s, 'Outside'); d.text(s, 1.05, 0.3, 0.3, 0.2, words(10), 18)         # 11 text box entirely outside the slide
s = d.slide(); head(s, 'Plain')                                                         # 12 nothing unusual
d.save('p22-geometry.pptx')

# =====================================================================================================
# p22-bleed-everywhere: every slide carries decorative bleed (accent bar off the top-left, footer band off three edges),
# one also a rotated rule inside the slide. Nothing overflows.
# =====================================================================================================
d = Deck(12192000, 6858000)
W, H = d.W, d.H
band = png((60, 50, 40))
for k, t in enumerate(['Who we are', 'What we do', 'Where we work', 'How we work', 'Who we serve']):
    s = d.slide()
    d.rect(s, -0.025, -0.03, 0.2, 0.07)                                                         # accent bar off the top-left corner
    s.shapes.add_picture(io.BytesIO(band), Emu(-100000), Emu(H - 300000), Emu(W + 200000), Emu(400000))   # footer band off three edges
    d.title(s, t); d.text(s, 0.08, 0.30, 0.8, 0.4, words(20), 20)
    if k == 1:
        r = d.rect(s, 0, 0, 0.7 * H / W, 0.004 * W / H, rot=90); r.left = Emu(int(W * .04 - H * .35)); r.top = Emu(int(H * .5 - W * .002))
d.save('p22-bleed-everywhere.pptx')
