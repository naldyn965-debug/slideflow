"""Phase 2.1 fixture: a dark template modelled on the real-world findings.
15 slides, #211D19 background, white text, full-slide pictures on slides 1 and 15, bleed/background shapes,
and the several ways PowerPoint-family tools express "white text on dark" (explicit colour, colour-map override,
shape style font colour, semi-transparent cards). Ground truth is in tests/e2e.mjs / tests/analysis.mjs."""
import io, struct, zlib, re, zipfile
from pptx import Presentation
from pptx.util import Inches, Pt, Emu
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE
from lxml import etree

A = 'http://schemas.openxmlformats.org/drawingml/2006/main'
P = 'http://schemas.openxmlformats.org/presentationml/2006/main'
def png(w=40, h=24, rgb=(60, 50, 40)):
    raw = b''.join(b'\x00' + bytes(rgb) * w for _ in range(h))
    def chunk(t, d): return struct.pack('>I', len(d)) + t + d + struct.pack('>I', zlib.crc32(t + d) & 0xffffffff)
    return b'\x89PNG\r\n\x1a\n' + chunk(b'IHDR', struct.pack('>IIBBBBB', w, h, 8, 2, 0, 0, 0)) + chunk(b'IDAT', zlib.compress(raw)) + chunk(b'IEND', b'')
PNG = png()
W, H = 12192000, 6858000
BG = '211D19'

p = Presentation(); p.slide_width, p.slide_height = Emu(W), Emu(H)
blank = p.slide_layouts[6]

def dark_bg(s):
    s.background.fill.solid(); s.background.fill.fore_color.rgb = RGBColor.from_string(BG)
def text(s, fx, fy, fw, fh, t, pt, color='FFFFFF', bold=False, font='Inter', raw_color=True):
    tb = s.shapes.add_textbox(Emu(int(W*fx)), Emu(int(H*fy)), Emu(int(W*fw)), Emu(int(H*fh))); tf = tb.text_frame; tf.word_wrap = True; tf.text = t
    for r in tf.paragraphs[0].runs:
        r.font.size = Pt(pt); r.font.bold = bold; r.font.name = font
        if raw_color: r.font.color.rgb = RGBColor.from_string(color)
    return tb
def picture(s, x, y, w, h): return s.shapes.add_picture(io.BytesIO(PNG), Emu(x), Emu(y), Emu(w), Emu(h))
def rect(s, x, y, w, h, hexv, alpha=None):
    r = s.shapes.add_shape(MSO_SHAPE.RECTANGLE, Emu(x), Emu(y), Emu(w), Emu(h)); r.fill.solid(); r.fill.fore_color.rgb = RGBColor.from_string(hexv); r.line.fill.background()
    if alpha is not None:
        clr = r._element.spPr.find(f'{{{A}}}solidFill/{{{A}}}srgbClr'); a = etree.SubElement(clr, f'{{{A}}}alpha'); a.set('val', str(int(alpha * 1000)))
    return r
def bleed_bar(s):  # decorative accent bleeding off the left/top edges
    rect(s, -300000, -200000, 2400000, 500000, 'C8873A')

# 1: cover, full-slide picture, title + subtitle over it
s = p.slides.add_slide(blank); dark_bg(s); picture(s, 0, 0, W, H)
text(s, 0.08, 0.35, 0.7, 0.2, 'Harvest Studio 2026', 48, bold=True); text(s, 0.08, 0.58, 0.6, 0.08, 'Brand and product overview', 22)
# 2-3: explicit white text, bleed bar
for k, (t, b) in enumerate([('Who we are', 'A small team building tools for growers, with a focus on clear data and honest packaging.'), ('What we do', 'We design, test and ship products that make daily farm work simpler for every crew.')]):
    s = p.slides.add_slide(blank); dark_bg(s); bleed_bar(s); text(s, 0.08, 0.12, 0.84, 0.12, t, 36, bold=True); text(s, 0.08, 0.30, 0.8, 0.4, b, 20)
# 4: section divider (title only), bleed bar
s = p.slides.add_slide(blank); dark_bg(s); bleed_bar(s); text(s, 0.08, 0.40, 0.8, 0.15, 'Part two: Products', 44, bold=True)
# 5: semi-transparent white card behind white text
s = p.slides.add_slide(blank); dark_bg(s); text(s, 0.08, 0.10, 0.84, 0.12, 'Highlights', 36, bold=True)
rect(s, int(W*.08), int(H*.30), int(W*.84), int(H*.5), 'FFFFFF', alpha=8)
text(s, 0.10, 0.34, 0.8, 0.4, 'Customer retention reached ninety two percent across all regions in the last full year of trading.', 20)
# 6: white text over a dark rectangle that is a separate shape
s = p.slides.add_slide(blank); dark_bg(s); text(s, 0.08, 0.10, 0.84, 0.12, 'Pricing', 36, bold=True)
rect(s, int(W*.08), int(H*.3), int(W*.4), int(H*.5), '3A332D'); text(s, 0.10, 0.34, 0.36, 0.3, 'Starter plan for single farms and co-ops with a shared account.', 20)
# 7: dark text on a light card over a dark slide (legitimate, must NOT be flagged)
s = p.slides.add_slide(blank); dark_bg(s); text(s, 0.08, 0.10, 0.84, 0.12, 'Quote', 36, bold=True)
rect(s, int(W*.08), int(H*.3), int(W*.5), int(H*.4), 'F3EDE4'); text(s, 0.10, 0.34, 0.46, 0.3, 'Good tools disappear into the work.', 28, color='211D19')
# 8: genuine low contrast (must STILL be flagged): grey 3B3733 text on the dark background
s = p.slides.add_slide(blank); dark_bg(s); text(s, 0.08, 0.10, 0.84, 0.12, 'Footnotes', 36, bold=True)
text(s, 0.08, 0.30, 0.8, 0.3, 'These notes are almost invisible against the background and need attention.', 20, color='3B3733')
# 9: white text over a content picture (contrast not evaluable)
s = p.slides.add_slide(blank); dark_bg(s); picture(s, int(W*.55), -100000, int(W*.45)+100000, H+200000); text(s, 0.08, 0.12, 0.38, 0.2, 'Field notes', 36, bold=True)
text(s, 0.08, 0.35, 0.38, 0.4, 'Observations collected from twelve partner farms during the spring season.', 18)
# 10: photo statement slide (full-slide picture, short text, mid-deck)
s = p.slides.add_slide(blank); dark_bg(s); picture(s, 0, 0, W, H); text(s, 0.08, 0.4, 0.8, 0.15, 'Grown with care', 44, bold=True)
# 11: long text over full-slide picture (text slide with a background image)
s = p.slides.add_slide(blank); dark_bg(s); picture(s, 0, 0, W, H); text(s, 0.08, 0.10, 0.84, 0.12, 'Our principles', 36, bold=True)
text(s, 0.08, 0.28, 0.84, 0.5, ' '.join(['We keep every decision tied to a measurable outcome for the people who use the product every day.'] * 3), 18)
# 12-13: plain text slides
for t in ('Roadmap', 'Team'):
    s = p.slides.add_slide(blank); dark_bg(s); bleed_bar(s); text(s, 0.08, 0.10, 0.84, 0.12, t, 36, bold=True)
    picture(s, -100000, H - 300000, W + 200000, 400000)  # decorative footer band bleeding off three edges
    text(s, 0.08, 0.30, 0.8, 0.4, 'Quarterly milestones are reviewed with partners before they are committed to the plan.', 20)
# 14: contact slide with wording
s = p.slides.add_slide(blank); dark_bg(s); text(s, 0.08, 0.10, 0.84, 0.12, 'Contact us', 36, bold=True); text(s, 0.08, 0.3, 0.8, 0.2, 'hello@example.com', 24)
# 15: closing, full-slide picture, NO closing wording, short text
s = p.slides.add_slide(blank); dark_bg(s); picture(s, 0, 0, W, H); text(s, 0.08, 0.42, 0.8, 0.15, 'Harvest Studio', 44, bold=True)
p.save('/tmp/P21a.pptx')

# ---- variant B: dark mode expressed through the colour map (what PowerPoint writes for dark background styles) ----
def ovr(src, dst, keep_bg=False):
    zin = zipfile.ZipFile(src); zout = zipfile.ZipFile(dst, 'w', zipfile.ZIP_DEFLATED)
    for n in zin.namelist():
        d = zin.read(n)
        if re.fullmatch(r'ppt/slides/slide\d+\.xml', n):
            x = d.decode('utf8')
            # drop explicit text colours so text follows tx1, and drop the explicit bg so bg1 decides
            x = re.sub(r'<a:solidFill><a:srgbClr val="(FFFFFF)"/></a:solidFill>(?=<a:latin)', '', x)
            if not keep_bg: x = re.sub(r'<p:bg>.*?</p:bg>', '<p:bg><p:bgRef idx="1001"><a:schemeClr val="bg1"/></p:bgRef></p:bg>', x, flags=re.S)
            x = x.replace('<p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr>', '<p:clrMapOvr><a:overrideClrMapping bg1="dk1" tx1="lt1" bg2="dk2" tx2="lt2" accent1="accent1" accent2="accent2" accent3="accent3" accent4="accent4" accent5="accent5" accent6="accent6" hlink="hlink" folHlink="folHlink"/></p:clrMapOvr>')
            d = x.encode('utf8')
        if n == 'ppt/theme/theme1.xml' and not keep_bg:
            x = d.decode('utf8'); x = re.sub(r'(<a:dk1>)<a:sysClr[^>]*/>', rf'\g<1><a:srgbClr val="{BG}"/>', x); d = x.encode('utf8')
        zout.writestr(n, d)
    zout.close()
ovr('/tmp/P21a.pptx', '/tmp/P21b.pptx')
ovr('/tmp/P21a.pptx', '/tmp/P21c.pptx', keep_bg=True)
import shutil; shutil.copy('/tmp/P21a.pptx', 'fixtures/p21-dark-explicit.pptx'); shutil.copy('/tmp/P21b.pptx', 'fixtures/p21-dark-clrmap.pptx'); shutil.copy('/tmp/P21c.pptx', 'fixtures/p21-dark-clrmap-explicit-bg.pptx')
print('p21 fixtures ok')

# ---------- second fixture: text colour from shape style, and a layout-level colour-map override ----------
p = Presentation(); p.slide_width, p.slide_height = Emu(W), Emu(H)
L = p.slide_layouts
def shape_text(s, hexv, t):
    sh = s.shapes.add_shape(MSO_SHAPE.RECTANGLE, Emu(int(W*.1)), Emu(int(H*.3)), Emu(int(W*.8)), Emu(int(H*.3)))
    sh.fill.solid(); sh.fill.fore_color.rgb = RGBColor.from_string(hexv); sh.line.fill.background()
    sh.text_frame.text = t
    for r in sh.text_frame.paragraphs[0].runs: r.font.size = Pt(24)   # colour deliberately NOT set: comes from <p:style><a:fontRef> (white)
s = p.slides.add_slide(L[6]); dark_bg(s); shape_text(s, '3A332D', 'White by style on a dark panel')
s = p.slides.add_slide(L[6]); dark_bg(s); shape_text(s, 'F3EDE4', 'White by style on a light panel, genuinely unreadable')
s = p.slides.add_slide(L[5]); dark_bg(s); s.shapes.title.text = 'Title inherits tx1 through the layout override'
s = p.slides.add_slide(L[6]); dark_bg(s); text(s, 0.10, 0.2, 0.5, 0.1, 'Headline', 36, bold=True)
text(s, 0.75, 0.5, 0.4, 0.1, 'This box runs off the right edge of the slide', 20)   # real problem: text off the slide
s = p.slides.add_slide(L[5]); lay_part = str(s.slide_layout.part.partname)  # keep the layout-override slide last-but-one: see index below
s.shapes.title.text = 'Layout override, second slide'
p.save('/tmp/P21d0.pptx')
zin = zipfile.ZipFile('/tmp/P21d0.pptx'); zout = zipfile.ZipFile('fixtures/p21-style-layout-map.pptx', 'w', zipfile.ZIP_DEFLATED)
for n in zin.namelist():
    d = zin.read(n)
    if '/' + n == lay_part:
        x = d.decode('utf8')
        ov = '<p:clrMapOvr><a:overrideClrMapping bg1="dk1" tx1="lt1" bg2="dk2" tx2="lt2" accent1="accent1" accent2="accent2" accent3="accent3" accent4="accent4" accent5="accent5" accent6="accent6" hlink="hlink" folHlink="folHlink"/></p:clrMapOvr>'
        assert '<p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr>' in x
        d = x.replace('<p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr>', ov).encode('utf8')
    zout.writestr(n, d)
zout.close()
print('p21 style/layout fixture ok')
