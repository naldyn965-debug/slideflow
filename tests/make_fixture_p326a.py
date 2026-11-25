"""Fixture for the Phase 3.2.6-A style edit-state tests: tests/fixtures/p326a-styles.pptx.
One 13.333 x 7.5 in slide. Run: python3 tests/make_fixture_p326a.py   (needs python-pptx, lxml)
Shapes (by name):
  solid    rectangle, solid fill FF0000, solid outline 0000FF 2 pt, one run "Plain" (Arial 20 pt, colour 112233)
  nofill   rectangle, no fill, no outline, one run "No fill" (18 pt)
  grad     rectangle, gradient fill, solid outline 00AA00, one run
  mixed    text box, one paragraph: bold red 18 pt / italic blue 24 pt / plain 20 pt / a hyperlink run
  scaled   text box, normAutofit fontScale=62500, one 32 pt run (colour 222222)
  emptyp   text box with one paragraph that has no run, <a:endParaRPr sz="2800">
  conn     straight connector, outline AA00AA 1.5 pt
  pic      picture (not styleable)
  tbl      table (not styleable)
  grp      group holding the rectangle grpchild (the group is not styleable, the child is)
"""
import io, os, struct, zlib
from pptx import Presentation
from pptx.util import Emu, Pt
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE, MSO_CONNECTOR
from pptx.enum.text import MSO_AUTO_SIZE
from lxml import etree

A = '{http://schemas.openxmlformats.org/drawingml/2006/main}'
PT = 12700
prs = Presentation(); prs.slide_width = Emu(12192000); prs.slide_height = Emu(6858000)
s = prs.slides.add_slide(prs.slide_layouts[6])

def rect(name, x, y, w, h, shapes=None):
    sh = (shapes or s.shapes).add_shape(MSO_SHAPE.RECTANGLE, Emu(x * PT), Emu(y * PT), Emu(w * PT), Emu(h * PT)); sh.name = name
    sh.shadow.inherit = False
    return sh

def run(p, text, size=20, rgb=None, bold=None, italic=None, font=None):
    r = p.add_run(); r.text = text; r.font.size = Pt(size)
    if rgb: r.font.color.rgb = RGBColor.from_string(rgb)
    if bold is not None: r.font.bold = bold
    if italic is not None: r.font.italic = italic
    if font: r.font.name = font
    return r

def box(name, x, y, w, h):
    t = s.shapes.add_textbox(Emu(x * PT), Emu(y * PT), Emu(w * PT), Emu(h * PT)); t.name = name
    t.text_frame.word_wrap = True; t.text_frame.auto_size = MSO_AUTO_SIZE.NONE
    return t

sh = rect('solid', 40, 20, 300, 70); sh.fill.solid(); sh.fill.fore_color.rgb = RGBColor.from_string('FF0000')
sh.line.color.rgb = RGBColor.from_string('0000FF'); sh.line.width = Pt(2)
run(sh.text_frame.paragraphs[0], 'Plain', 20, '112233', font='Arial')
sh = rect('nofill', 40, 110, 300, 70); sh.fill.background(); sh.line.fill.background()
run(sh.text_frame.paragraphs[0], 'No fill', 18)
sh = rect('grad', 40, 200, 300, 70); sh.fill.gradient(); sh.line.color.rgb = RGBColor.from_string('00AA00'); sh.line.width = Pt(1)
run(sh.text_frame.paragraphs[0], 'Gradient', 18)
t = box('mixed', 40, 290, 420, 60); p = t.text_frame.paragraphs[0]
run(p, 'Bold red ', 18, 'FF0000', bold=True); run(p, 'italic blue ', 24, '0000FF', italic=True); run(p, 'plain ', 20)
r = run(p, 'link', 20); r.hyperlink.address = 'https://example.com'
t = box('scaled', 500, 20, 280, 100); tf = t.text_frame
bp = tf._txBody.find(A + 'bodyPr')
for c in list(bp): bp.remove(c)
n = etree.SubElement(bp, A + 'normAutofit'); n.set('fontScale', '62500'); n.set('lnSpcReduction', '20000')
run(tf.paragraphs[0], 'Stored fontScale.', 32, '222222')
t = box('emptyp', 500, 150, 280, 60); p = t.text_frame.paragraphs[0]; p._p.get_or_add_endParaRPr().set('sz', '2800')
c = s.shapes.add_connector(MSO_CONNECTOR.STRAIGHT, Emu(500 * PT), Emu(240 * PT), Emu(800 * PT), Emu(300 * PT)); c.name = 'conn'
c.line.color.rgb = RGBColor.from_string('AA00AA'); c.line.width = Pt(1.5)
def png(w=4, h=4):
    raw = b''.join(b'\x00' + b'\xff\x00\x00' * w for _ in range(h))
    chunk = lambda k, d: struct.pack('>I', len(d)) + k + d + struct.pack('>I', zlib.crc32(k + d) & 0xffffffff)
    return b'\x89PNG\r\n\x1a\n' + chunk(b'IHDR', struct.pack('>IIBBBBB', w, h, 8, 2, 0, 0, 0)) + chunk(b'IDAT', zlib.compress(raw)) + chunk(b'IEND', b'')
pic = s.shapes.add_picture(io.BytesIO(png()), Emu(500 * PT), Emu(320 * PT), Emu(80 * PT), Emu(80 * PT)); pic.name = 'pic'
tb = s.shapes.add_table(2, 2, Emu(600 * PT), Emu(320 * PT), Emu(300 * PT), Emu(80 * PT)); tb.name = 'tbl'
g = s.shapes.add_group_shape(); g.name = 'grp'
ch = rect('grpchild', 40, 400, 200, 60, g.shapes); ch.fill.solid(); ch.fill.fore_color.rgb = RGBColor.from_string('00FF00'); ch.line.color.rgb = RGBColor.from_string('333333')
run(ch.text_frame.paragraphs[0], 'Group child', 18, 'FFFFFF')
out = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'fixtures', 'p326a-styles.pptx'); prs.save(out); print('wrote', out)
