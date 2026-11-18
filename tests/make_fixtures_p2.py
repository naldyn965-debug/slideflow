"""Phase 2 fixtures with known ground truth. Built with python-pptx (A,B,C,E,F) and pptxgenjs (D, see make_fixture_pptxgenjs.mjs)."""
import io, re, struct, zlib, zipfile, copy
from pptx import Presentation
from pptx.util import Inches, Pt, Emu
from pptx.dml.color import RGBColor
from pptx.enum.dml import MSO_THEME_COLOR
from pptx.enum.shapes import MSO_SHAPE
from pptx.chart.data import CategoryChartData
from pptx.enum.chart import XL_CHART_TYPE
from lxml import etree

out = 'fixtures/'
def png(w=40, h=24, rgb=(120, 160, 200)):
    raw = b''.join(b'\x00' + bytes(rgb) * w for _ in range(h))
    def chunk(t, d): return struct.pack('>I', len(d)) + t + d + struct.pack('>I', zlib.crc32(t + d) & 0xffffffff)
    return b'\x89PNG\r\n\x1a\n' + chunk(b'IHDR', struct.pack('>IIBBBBB', w, h, 8, 2, 0, 0, 0)) + chunk(b'IDAT', zlib.compress(raw)) + chunk(b'IEND', b'')
PNG = png()

def retheme(src, dst, colors, major, minor):
    zin = zipfile.ZipFile(src); zout = zipfile.ZipFile(dst, 'w', zipfile.ZIP_DEFLATED)
    for n in zin.namelist():
        d = zin.read(n)
        if n == 'ppt/theme/theme1.xml':
            x = d.decode('utf8')
            for slot, hexv in colors.items():
                x = re.sub(rf'(<a:{slot}>)<a:srgbClr val="[0-9A-Fa-f]{{6}}"/>', rf'\g<1><a:srgbClr val="{hexv}"/>', x)
            x = re.sub(r'(<a:majorFont><a:latin typeface=")[^"]*', rf'\g<1>{major}', x)
            x = re.sub(r'(<a:minorFont><a:latin typeface=")[^"]*', rf'\g<1>{minor}', x)
            d = x.encode('utf8')
        zout.writestr(n, d)
    zout.close()

def logo(s, x=8.6, y=0.15):
    s.shapes.add_picture(io.BytesIO(PNG), Inches(x), Inches(y), Inches(1.0), Inches(0.6))

# ---------- A: 4:3 corporate deck, themed, layout-driven ----------
p = Presentation()
L = p.slide_layouts
s = p.slides.add_slide(L[0]); s.shapes.title.text = 'Acme Corp 2026 Strategy'; s.placeholders[1].text = 'Board briefing'
s = p.slides.add_slide(L[2]); s.shapes.title.text = 'Part 1: Where we are'; s.placeholders[1].text = 'Market and performance'
bar = s.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(0), Inches(0), Inches(10), Inches(0.3)); bar.fill.solid(); bar.fill.fore_color.theme_color = MSO_THEME_COLOR.ACCENT_1; bar.line.fill.background()
words = 'Revenue grew twelve percent while churn fell to a record low across every region we serve today'
for k, extra in enumerate(['', ' and enterprise accounts doubled', ' with margins improving steadily']):
    s = p.slides.add_slide(L[1]); s.shapes.title.text = f'Highlights {k+1}'
    tf = s.placeholders[1].text_frame; tf.text = words + extra
    tf.add_paragraph().text = 'Customer satisfaction remains above target for the fourth quarter'
    logo(s)
    if k == 0:  # explicit sizes on slide 3
        for para in tf.paragraphs:
            for r in para.runs: r.font.size = Pt(20)
        for r in s.shapes.title.text_frame.paragraphs[0].runs: r.font.size = Pt(40); r.font.bold = True
    if k == 2:  # 9pt footnote
        fb = s.shapes.add_textbox(Inches(0.5), Inches(6.9), Inches(6), Inches(0.3)); fb.text_frame.text = 'Source: internal finance data, unaudited'
        fb.text_frame.paragraphs[0].runs[0].font.size = Pt(9)
s = p.slides.add_slide(L[5]); s.shapes.title.text = 'Risks'  # dark, low contrast
s.background.fill.solid(); s.background.fill.fore_color.rgb = RGBColor(0x1A, 0x1A, 0x1A)
tb = s.shapes.add_textbox(Inches(0.8), Inches(2), Inches(8), Inches(2)); tb.text_frame.text = 'Supply chain exposure remains elevated in two regions and we are diversifying suppliers to reduce the risk before next year'
for r in tb.text_frame.paragraphs[0].runs: r.font.size = Pt(24); r.font.color.rgb = RGBColor(0x33, 0x33, 0x33)
for r in s.shapes.title.text_frame.paragraphs[0].runs: r.font.color.rgb = RGBColor(0xFF, 0xFF, 0xFF)
s = p.slides.add_slide(L[6]); s.shapes.add_picture(io.BytesIO(PNG), Inches(1), Inches(0.8), Inches(6.0), Inches(4.0)); logo(s)
cap = s.shapes.add_textbox(Inches(1), Inches(5.2), Inches(6), Inches(0.5)); cap.text_frame.text = 'Our new Berlin office'
s = p.slides.add_slide(L[4]); s.shapes.title.text = 'Build versus buy'
for idx, t in {1: 'Build', 2: 'Full control over the roadmap and data, slower to launch and costly to staff', 3: 'Buy', 4: 'Faster launch with vendor support, less control and recurring licence fees'}.items(): s.placeholders[idx].text = t
s = p.slides.add_slide(L[5]); s.shapes.title.text = 'Revenue by region'
cd = CategoryChartData(); cd.categories = ['EMEA', 'APAC', 'AMER']; cd.add_series('2026', (4, 5, 6))
s.shapes.add_chart(XL_CHART_TYPE.COLUMN_CLUSTERED, Inches(1), Inches(1.5), Inches(8), Inches(5), cd)
s = p.slides.add_slide(L[5]); s.shapes.title.text = 'Pricing'
t = s.shapes.add_table(3, 3, Inches(1), Inches(2), Inches(8), Inches(2)).table
for r in range(3):
    for c in range(3): t.cell(r, c).text = f'R{r}C{c}'
s = p.slides.add_slide(L[0]); s.shapes.title.text = 'Thank you'; s.placeholders[1].text = 'Questions?'
p.save('/tmp/A0.pptx')
retheme('/tmp/A0.pptx', out + 'p2-corporate.pptx', {'accent1': 'C0392B', 'accent2': '2C3E50', 'accent3': '16A085', 'accent4': 'F39C12', 'dk2': '2C3E50', 'lt2': 'ECF0F1'}, 'Georgia', 'Verdana')

# ---------- B: 16:9, fully explicit geometry, text boxes only (ground-truth margins) ----------
W, H = 12192000, 6858000
p = Presentation(); p.slide_width, p.slide_height = Emu(W), Emu(H)
def box(s, fx, fy, fw, fh, text, pt, rgb='111111', bold=False, font='Arial', align=None):
    tb = s.shapes.add_textbox(Emu(int(W*fx)), Emu(int(H*fy)), Emu(int(W*fw)), Emu(int(H*fh))); tf = tb.text_frame; tf.word_wrap = True
    tf.text = text
    for para in tf.paragraphs:
        if align: para.alignment = align
        for r in para.runs: r.font.size = Pt(pt); r.font.bold = bold; r.font.name = font; r.font.color.rgb = RGBColor.from_string(rgb)
for title, body in [('Our approach', 'We start from customer problems and work backwards to the smallest useful product.'), ('How we work', 'Weekly releases, shared metrics and one owner for every decision we take together.'), ('What we found', 'Teams that measure outcomes ship twice as often and report fewer incidents each quarter.'), ('Contact us', 'hello@example.com')]:
    s = p.slides.add_slide(p.slide_layouts[6])
    box(s, 0.08, 0.10, 0.84, 0.12, title, 36, bold=True); box(s, 0.08, 0.30, 0.84, 0.55, body, 20)
p.save(out + 'p2-explicit.pptx')

# ---------- C: Arabic RTL with explicit cs font ----------
p = Presentation(); p.slide_width, p.slide_height = Emu(W), Emu(H)
def arabic(s, y, text, pt, font='Cairo'):
    tb = s.shapes.add_textbox(Emu(int(W*0.1)), Emu(int(H*y)), Emu(int(W*0.8)), Emu(int(H*0.15))); tb.text_frame.text = text
    pr = tb.text_frame.paragraphs[0]; pr.alignment = 3  # right
    for r in pr.runs:
        r.font.size = Pt(pt); rPr = r._r.get_or_add_rPr(); cs = etree.SubElement(rPr, '{http://schemas.openxmlformats.org/drawingml/2006/main}cs'); cs.set('typeface', font)
for t1, t2 in [('مرحبا بكم في عرضنا', 'الزراعة الذكية للمستقبل'), ('المحتوى الرئيسي', 'نناقش اليوم ثلاثة محاور أساسية تخص تطوير المنتج وخدمة العملاء وتحسين الجودة'), ('شكراً لكم', 'هل لديكم أسئلة؟')]:
    s = p.slides.add_slide(p.slide_layouts[6]); arabic(s, 0.15, t1, 40); arabic(s, 0.4, t2, 24)
p.save(out + 'p2-arabic.pptx')

# ---------- E: theme part removed (degenerate) ----------
zin = zipfile.ZipFile(out + 'p2-explicit.pptx'); zout = zipfile.ZipFile(out + 'p2-no-theme.pptx', 'w', zipfile.ZIP_DEFLATED)
for n in zin.namelist():
    if n.startswith('ppt/theme/'): continue
    d = zin.read(n)
    if n == '[Content_Types].xml': d = re.sub(rb'<Override PartName="/ppt/theme/[^>]*/>', b'', d)
    if n.endswith('.rels'): d = re.sub(rb'<Relationship [^>]*relationships/theme"[^>]*/>', b'', d)
    zout.writestr(n, d)
zout.close()

# ---------- F: picture-only single slide ----------
p = Presentation(); p.slide_width, p.slide_height = Emu(W), Emu(H)
s = p.slides.add_slide(p.slide_layouts[6]); s.shapes.add_picture(io.BytesIO(PNG), 0, 0, Emu(W), Emu(H))
p.save(out + 'p2-image-only.pptx')
print('p2 fixtures ok')
