"""Phase 3.2.1 fixture: ground truth for element selection (tests/selection.mjs).
Run from the repo root:  python3 tests/make_fixture_p321.py   (needs python-pptx; reuses the raw-XML builders of make_fixture_p31.py).

p321-selection.pptx   13.33 x 7.5 in (960 x 540 pt). Every frame below is in POINTS, so a test at 960 px wide can click exact pixels.
  Layout and master carry decoration that covers the slide (a full-slide wash on the layout, a strip on the master): selecting must ignore it,
  and clicking it must count as empty space.
  Slide 1  back-rect (60,60,300x200) under front-rect (200,140,300x200): overlap by z-order
           under-rect (500,40,260x240) with circle (520,60,200x200) on top: the frame corner of the circle is rect, not circle
           h-line (60,400,300x0, 1 pt) and diag-line (400,380,200x120, 2 pt): hairlines; the diagonal's bounding box must not steal clicks
           outline-only (640,320,240x160, no fill, 3 pt border) with inner-fill (680,360,120x80) on top
           rot-rect (780,70,160x60, rotated 45)   textbox (60,300,130x40, no fill, "Label")
           grp (380,20,110x110) with g-child-a (rect) and g-child-b (ellipse)   pic (780,190,100x100, ellipse crop)
           hidden-rect (60,470,100x60, hidden)   table (200,440,180x60)
  Slide 2  s2-rect (100,100,200x100) and s2-text; its id 2 is the same number as back-rect's on slide 1 (ids are qualified by the slide part)
  Slide 3  dup-bottom and dup-top with the very same frame (100,100,300x200); flip-line (flipH, 500,100,200x150); rot-line (rotated 30, 100,420,300x0)
           hair-under (420,430,300x0) below card-over (500,420,100x40); card-under (420,470,100x40) below hair-over (420,490,300x0)
"""
import os, sys, tempfile
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from make_fixture_p31 import (X, put, png, sp, tbox, grp, cxn, para, run, SOLID, LINE, new_deck, blank, W, H, OUT)

P = lambda pt: int(round(pt * 12700))
NOFILL = '<a:noFill/>'
NOLINE = '<a:ln><a:noFill/></a:ln>'


def box(x, y, w, h): return P(x), P(y), P(w), P(h)


def make():
    p = new_deck()
    # ---- template decoration (inert for selection) ----
    master, layout = p.slide_master, p.slide_layouts[6]
    master.shapes._spTree.insert(2, X(sp(900, 'master-strip', *box(0, 510, 960, 30), fill=SOLID('D9D9D9'), ln=NOLINE)))
    layout.shapes._spTree.insert(2, X(sp(901, 'layout-wash', *box(0, 0, 960, 540), fill=SOLID('F7F7F7'), ln=NOLINE)))

    s1 = blank(p)
    put(s1,
        sp(2, 'back-rect', *box(60, 60, 300, 200), fill=SOLID('4472C4'), ln=NOLINE),
        sp(3, 'front-rect', *box(200, 140, 300, 200), fill=SOLID('ED7D31'), ln=NOLINE),
        sp(5, 'under-rect', *box(500, 40, 260, 240), fill=SOLID('C00000'), ln=NOLINE),
        sp(4, 'circle', *box(520, 60, 200, 200), prst='ellipse', fill=SOLID('70AD47'), ln=NOLINE),
        cxn(6, 'h-line', *box(60, 400, 300, 0), prst='line', ln=LINE('000000', 12700)),
        cxn(7, 'diag-line', *box(400, 380, 200, 120), prst='line', ln=LINE('7030A0', 25400)),
        sp(8, 'outline-only', *box(640, 320, 240, 160), fill=NOFILL, ln=LINE('1F3864', 38100)),
        sp(9, 'inner-fill', *box(680, 360, 120, 80), fill=SOLID('A6A6A6'), ln=NOLINE),
        sp(10, 'rot-rect', *box(780, 70, 160, 60), rot=45, fill=SOLID('7030A0'), ln=NOLINE),
        tbox(11, 'textbox', *box(60, 300, 130, 40), para(run('Label', 18))),
        grp(12, 'grp', *box(380, 20, 110, 110), (P(380), P(20), P(110), P(110)),
            sp(13, 'g-child-a', *box(380, 20, 60, 60), fill=SOLID('00B0F0'), ln=NOLINE),
            sp(14, 'g-child-b', *box(420, 60, 70, 70), prst='ellipse', fill=SOLID('FFC000'), ln=NOLINE)),
        sp(16, 'hidden-rect', *box(60, 470, 100, 60), fill=SOLID('FF0000'), ln=NOLINE))
    # hidden flag, picture (ellipse crop) and table go in through python-pptx, then get fixed ids / names
    for sh in s1.shapes._spTree.iter('{http://schemas.openxmlformats.org/presentationml/2006/main}cNvPr'):
        if sh.get('name') == 'hidden-rect': sh.set('hidden', '1')
    d = tempfile.mkdtemp(); fp = os.path.join(d, 'p.png'); open(fp, 'wb').write(png((200, 60, 60), 80, 80))
    pic = s1.shapes.add_picture(fp, *box(780, 190, 100, 100))
    pic._element.nvPicPr.cNvPr.set('id', '15'); pic._element.nvPicPr.cNvPr.set('name', 'pic')
    pic._element.spPr.find('{http://schemas.openxmlformats.org/drawingml/2006/main}prstGeom').set('prst', 'ellipse')
    tbl = s1.shapes.add_table(2, 2, *box(200, 440, 180, 60))
    tbl._element.nvGraphicFramePr.cNvPr.set('id', '17'); tbl._element.nvGraphicFramePr.cNvPr.set('name', 'table')

    s2 = blank(p)
    put(s2,
        sp(2, 's2-rect', *box(100, 100, 200, 100), fill=SOLID('2E75B6'), ln=NOLINE),
        tbox(3, 's2-text', *box(400, 100, 300, 50), para(run('Second slide', 24))))

    s3 = blank(p)
    put(s3,
        sp(2, 'dup-bottom', *box(100, 100, 300, 200), fill=SOLID('4472C4'), ln=NOLINE),
        sp(3, 'dup-top', *box(100, 100, 300, 200), fill=SOLID('ED7D31'), ln=NOLINE),
        cxn(4, 'flip-line', *box(500, 100, 200, 150), prst='line', fh=True, ln=LINE('C00000', 12700)),
        cxn(5, 'rot-line', *box(100, 420, 300, 0), prst='line', rot=30, ln=LINE('00B050', 12700)),
        cxn(6, 'hair-under', *box(420, 430, 300, 0), prst='line', ln=LINE('000000', 12700)),
        sp(7, 'card-over', *box(500, 420, 100, 40), fill=SOLID('A6A6A6'), ln=NOLINE),
        sp(8, 'card-under', *box(420, 470, 100, 40), fill=SOLID('BFBFBF'), ln=NOLINE),
        cxn(9, 'hair-over', *box(420, 490, 300, 0), prst='line', ln=LINE('C00000', 12700)))
    p.save(OUT + 'p321-selection.pptx'); print('wrote p321-selection.pptx')


if __name__ == '__main__':
    make()
