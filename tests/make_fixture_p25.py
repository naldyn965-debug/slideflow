"""Phase 2.5 / Step 2 fixtures: Enhance rule coverage. Ground truth is asserted in tests/analysis.mjs (section "Phase 2.5 Step 2").
Run from the repo root:  python3 tests/make_fixture_p25.py   (needs python-pptx; also run by `npm run fixtures`).

  p25-food-like.pptx           20 x 11.25 in (18288000 x 10287000 EMU, the size of the real FOOD deck): text boxes only, no placeholders, dark
                               #211D19 background, 21 pt white justified body, 80 pt bold titles on a yellow panel, brand name + typed page number
                               on slides 2-7, a 194 pt cover and a 165 pt closing. Everything is at or above every Enhance floor.
  p25-enh-sizes.pptx           16:9. One deck that trips all four size rules: body 12 pt, title 20 pt, subtitle 14 pt, footer 9 pt.
  p25-enh-title-ratio.pptx     16:9. Title 30 pt over 28 pt body: above the 28 pt floor but under 1.25 x body, so the ratio rule decides (35 pt).
  p25-enh-margins.pptx         16:9. Content frame at 2 / 2 / 2 / 1 % (left / top / right / bottom), all under the 5 / 5 / 5 / 4 floors.
  p25-enh-contrast-light.pptx  16:9. #BBBBBB body text on a white background (about 1.9:1): the text colour must become the theme's dk1.
  p25-enh-contrast-dark.pptx   16:9. #444444 body text on #211D19 (about 1.7:1): the text colour must become the theme's lt1.
  p25-enh-at-floors.pptx       16:9. Every value exactly ON its floor (body 16, title 28, subtitle 16, footer 12, margins 5 / 5 / 5 / 4, readable
                               colours): a floor is a minimum, so nothing may change.

All values are explicit runs and explicit geometry, so the expected tokens do not depend on python-pptx's default master.
"""
from lxml import etree
from pptx import Presentation
from pptx.util import Emu, Pt
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.text import PP_ALIGN

OUT = 'tests/fixtures/'
NS = 'xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"'
W169, H169 = 12192000, 6858000
LOREM = ('Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod tempor incididunt ut labore et dolore magna aliqua. '
         'Ut enim ad minim veniam, quis nostrud exercitation ullamco laboris nisi ut aliquip ex ea commodo consequat. '
         'Duis aute irure dolor in reprehenderit in voluptate velit esse cillum dolore eu fugiat nulla pariatur.')


def rgb(h): return RGBColor.from_string(h)


def style(tf, pt, color=None, bold=False, font=None, align=None):
    for p in tf.paragraphs:
        if align is not None: p.alignment = align
        for r in p.runs:
            r.font.size = Pt(pt); r.font.bold = bold
            if color: r.font.color.rgb = rgb(color)
            if font: r.font.name = font


def background(slide, hexv):
    slide.background.fill.solid(); slide.background.fill.fore_color.rgb = rgb(hexv)


def place(shape, W, H, x, y, w, h):
    shape.left, shape.top, shape.width, shape.height = Emu(round(W * x)), Emu(round(H * y)), Emu(round(W * w)), Emu(round(H * h))


def textbox(slide, W, H, x, y, w, h, text, pt, **kw):
    b = slide.shapes.add_textbox(Emu(round(W * x)), Emu(round(H * y)), Emu(round(W * w)), Emu(round(H * h)))
    b.text_frame.word_wrap = True; b.text_frame.text = text
    style(b.text_frame, pt, **kw)
    return b


def rect(slide, W, H, x, y, w, h, hexv):
    r = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Emu(round(W * x)), Emu(round(H * y)), Emu(round(W * w)), Emu(round(H * h)))
    r.fill.solid(); r.fill.fore_color.rgb = rgb(hexv); r.line.fill.background()
    return r


def footer(slide, W, H, text, pt):
    """A real footer placeholder (type ftr) with explicit geometry and size, so the footer role is assigned."""
    x, y, w, h = round(W * 0.05), round(H * 0.965), round(W * 0.30), round(H * 0.03)
    xml = (f'<p:sp {NS}><p:nvSpPr><p:cNvPr id="90" name="Footer Placeholder 90"/><p:cNvSpPr><a:spLocks noGrp="1"/></p:cNvSpPr>'
           f'<p:nvPr><p:ph type="ftr" sz="quarter" idx="11"/></p:nvPr></p:nvSpPr><p:spPr><a:xfrm><a:off x="{x}" y="{y}"/><a:ext cx="{w}" cy="{h}"/></a:xfrm></p:spPr>'
           f'<p:txBody><a:bodyPr/><a:lstStyle/><a:p><a:r><a:rPr lang="en-US" sz="{int(pt * 100)}"/><a:t>{text}</a:t></a:r></a:p></p:txBody></p:sp>')
    slide.shapes._spTree.append(etree.fromstring(xml))


def enh_deck(name, title_pt, body_pt, sub_pt=None, footer_pt=None, margins=(8, 8, 8, 8), body_color='222222', title_color='111111', bg=None, slides=3):
    """16:9, `slides` identical slides built on the Title Slide layout. Title and subtitle are real placeholders; the body is a plain text box.
    margins = (left, top, right, bottom) in % of the slide: the title sits at (left, top); the body box ends at (100 - right, 100 - bottom).
    That makes the template-wide content margins exactly `margins`, because the other shapes sit inside that frame."""
    W, H = W169, H169
    l, t, r, b = margins
    p = Presentation(); p.slide_width, p.slide_height = Emu(W), Emu(H)
    for i in range(slides):
        s = p.slides.add_slide(p.slide_layouts[0])
        if bg: background(s, bg)
        s.shapes.title.text = f'Heading {i + 1}'
        place(s.shapes.title, W, H, l / 100, t / 100, 0.40, 0.10)
        style(s.shapes.title.text_frame, title_pt, color=title_color, bold=True, align=PP_ALIGN.LEFT)
        sub = s.placeholders[1]
        if sub_pt is None:
            sub._element.getparent().remove(sub._element)
        else:
            sub.text = 'A short subtitle line'
            place(sub, W, H, l / 100, (t + 12) / 100, 0.40, 0.08)
            style(sub.text_frame, sub_pt, color=body_color, align=PP_ALIGN.LEFT)
        bw, bh = 0.40, 0.30
        textbox(s, W, H, (100 - r) / 100 - bw, (100 - b) / 100 - bh, bw, bh, LOREM, body_pt, color=body_color, align=PP_ALIGN.LEFT)
        if footer_pt is not None: footer(s, W, H, 'Footer text', footer_pt)
    p.save(OUT + name); print('wrote', name)


def food_like():
    W, H = 18288000, 10287000                      # 20 x 11.25 in, like the real FOOD deck
    DARK, YEL = '211D19', 'FFFF00'
    p = Presentation(); p.slide_width, p.slide_height = Emu(W), Emu(H)
    blank = p.slide_layouts[6]
    titles = ['What is Food?', 'Food Groups', 'Fruits', 'Vegetables', 'Proteins', 'Grains']
    body = ' '.join(LOREM.split()[:60])

    s = p.slides.add_slide(blank); background(s, DARK)                                  # 1 cover (brand name at a different place than on 2-7)
    textbox(s, W, H, 0.295, 0.321, 0.411, 0.324, 'FOOD', 194, color='FFFFFF', bold=True, font='Inter Bold', align=PP_ALIGN.CENTER)
    textbox(s, W, H, 0.342, 0.200, 0.317, 0.059, 'The Wonderful World of', 35, color='FFFFFF', bold=True, font='Inter Bold', align=PP_ALIGN.CENTER)
    textbox(s, W, H, 0.319, 0.608, 0.362, 0.095, 'Let\u2019s explore different kinds of food and why they matter', 28, color='FFFFFF', font='Inter', align=PP_ALIGN.CENTER)
    textbox(s, W, H, 0.383, 0.788, 0.235, 0.035, 'www.example.com', 21, color='FFFFFF', bold=True, font='Inter Bold', align=PP_ALIGN.CENTER)
    textbox(s, W, H, 0.464, 0.186, 0.095, 0.035, 'Acme & Co.', 21, color='FFFFFF', bold=True, font='Inter Bold', align=PP_ALIGN.LEFT)

    for k, t in enumerate(titles):                                                      # 2-7 text slides
        s = p.slides.add_slide(blank); background(s, DARK)
        rect(s, W, H, 0.0, 0.90, 0.056, 0.10, YEL)
        rect(s, W, H, 0.056, 0.04, 0.60, 0.16, YEL)
        textbox(s, W, H, 0.874, 0.048, 0.095, 0.035, 'Acme & Co.', 21, color='FFFFFF', bold=True, font='Inter Bold', align=PP_ALIGN.LEFT)
        textbox(s, W, H, 0.0, 0.912, 0.056, 0.069, f'{k + 2:02d}', 42, color=DARK, bold=True, font='Inter Bold', align=PP_ALIGN.CENTER)
        textbox(s, W, H, 0.066, 0.051, 0.58, 0.133, t, 80, color=DARK, bold=True, font='Inter Bold', align=PP_ALIGN.LEFT)
        textbox(s, W, H, 0.056, 0.31, 0.818, 0.545, body, 21, color='FFFFFF', font='Inter', align=PP_ALIGN.JUSTIFY)

    s = p.slides.add_slide(blank); background(s, DARK)                                  # 8 closing
    textbox(s, W, H, 0.20, 0.30, 0.60, 0.30, 'THANK YOU', 165, color='FFFFFF', bold=True, font='Inter Bold', align=PP_ALIGN.CENTER)
    textbox(s, W, H, 0.383, 0.788, 0.235, 0.035, 'www.example.com', 21, color='FFFFFF', bold=True, font='Inter Bold', align=PP_ALIGN.CENTER)
    p.save(OUT + 'p25-food-like.pptx'); print('wrote p25-food-like.pptx')


food_like()
enh_deck('p25-enh-sizes.pptx', title_pt=20, body_pt=12, sub_pt=14, footer_pt=9)
enh_deck('p25-enh-title-ratio.pptx', title_pt=30, body_pt=28)
enh_deck('p25-enh-margins.pptx', title_pt=36, body_pt=20, margins=(2, 2, 2, 1))
enh_deck('p25-enh-contrast-light.pptx', title_pt=36, body_pt=20, body_color='BBBBBB', title_color='000000', bg='FFFFFF')
enh_deck('p25-enh-contrast-dark.pptx', title_pt=36, body_pt=20, body_color='444444', title_color='FFFFFF', bg='211D19')
enh_deck('p25-enh-at-floors.pptx', title_pt=28, body_pt=16, sub_pt=16, footer_pt=12, margins=(5, 5, 5, 4))
