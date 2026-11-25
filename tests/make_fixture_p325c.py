"""Fixture for the Phase 3.2.5-C regression tests: tests/fixtures/p325c-fixes.pptx.
One 13.333 x 7.5 in slide (960 x 540 pt, so 1 pt = 1 px when drawn 960 px wide). Run: python3 tests/make_fixture_p325c.py   (needs python-pptx, lxml)
Shapes (by name):
  single   one paragraph, one run                              -> emptied paragraph, re-edit (fix 1)
  twopara  two paragraphs                                       -> an emptied paragraph must not block edits of the other (fix 1)
  bul      two bulleted paragraphs                              -> emptied bulleted paragraph (fix 1, bullets kept)
  rtl      one right-to-left Arabic paragraph                   -> emptied RTL paragraph (fix 1, RTL kept)
  bronly   text / break only / two breaks only / text           -> line-break-only paragraphs (fix 2)
  allbr    one paragraph that is only a line break              -> nothing to edit (fix 2)
  brlast   text, then a paragraph that is only a line break     -> the caret of a fresh edit goes into the last paragraph that can be typed into (fix 2)
  brtext   one paragraph: run, line break, run                  -> still editable (fix 2: only break-ONLY paragraphs are locked)
  shrink   normAutofit without a stored fontScale               -> autofit baseline (fix 3 / 4)
  scaled   normAutofit fontScale=62500 lnSpcReduction=20000     -> stored fontScale (fix 3 / 4)
"""
from pptx import Presentation
from pptx.util import Inches, Pt, Emu
from pptx.enum.text import MSO_AUTO_SIZE
from lxml import etree
import os

A = '{http://schemas.openxmlformats.org/drawingml/2006/main}'
prs = Presentation(); prs.slide_width = Emu(12192000); prs.slide_height = Emu(6858000)
s = prs.slides.add_slide(prs.slide_layouts[6])
PT = 12700

def box(name, x, y, w, h):
    t = s.shapes.add_textbox(Emu(x * PT), Emu(y * PT), Emu(w * PT), Emu(h * PT)); t.name = name
    t.text_frame.word_wrap = True; t.text_frame.auto_size = MSO_AUTO_SIZE.NONE
    return t

def run(p, text, size=24):
    r = p.add_run(); r.text = text; r.font.size = Pt(size); return r

t = box('single', 40, 20, 420, 50); run(t.text_frame.paragraphs[0], 'Delete me all', 28)
t = box('twopara', 40, 90, 420, 90); tf = t.text_frame; run(tf.paragraphs[0], 'First paragraph'); run(tf.add_paragraph(), 'Second paragraph')
t = box('bronly', 40, 200, 420, 150); tf = t.text_frame
run(tf.paragraphs[0], 'Line one'); tf.add_paragraph().add_line_break()
p = tf.add_paragraph(); p.add_line_break(); p.add_line_break(); run(tf.add_paragraph(), 'Line four')
t = box('allbr', 40, 370, 420, 40); t.text_frame.paragraphs[0].add_line_break()
t = box('brlast', 500, 480, 420, 50); run(t.text_frame.paragraphs[0], 'Last text'); t.text_frame.add_paragraph().add_line_break()
t = box('brtext', 40, 430, 420, 80); p = t.text_frame.paragraphs[0]; run(p, 'Line A'); p.add_line_break(); run(p, 'Line B')
t = box('shrink', 500, 20, 280, 100); tf = t.text_frame
bp = tf._txBody.find(A + 'bodyPr')
for c in list(bp): bp.remove(c)
etree.SubElement(bp, A + 'normAutofit')
run(tf.paragraphs[0], 'Autofit shrink box with a modest amount of text here')
t = box('scaled', 500, 150, 280, 100); tf = t.text_frame
bp = tf._txBody.find(A + 'bodyPr')
for c in list(bp): bp.remove(c)
n = etree.SubElement(bp, A + 'normAutofit'); n.set('fontScale', '62500'); n.set('lnSpcReduction', '20000')
run(tf.paragraphs[0], 'Stored fontScale 62.5 percent. A longer paragraph that PowerPoint shrank to fit the frame.', 32)
t = box('bul', 500, 280, 420, 100); tf = t.text_frame
for txt in ('First bullet', 'Second bullet'):
    p = tf.paragraphs[0] if txt.startswith('First') else tf.add_paragraph(); run(p, txt, 22)
    pPr = p._p.get_or_add_pPr(); pPr.set('marL', '342900'); pPr.set('indent', '-342900')
    etree.SubElement(pPr, A + 'buChar').set('char', '•')
t = box('rtl', 500, 400, 420, 60); p = t.text_frame.paragraphs[0]
pPr = p._p.get_or_add_pPr(); pPr.set('rtl', '1'); pPr.set('algn', 'r')
r = run(p, 'مرحبا بالعالم', 28); r._r.get_or_add_rPr().set('lang', 'ar-SA')
out = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'fixtures', 'p325c-fixes.pptx'); prs.save(out); print('wrote', out)
