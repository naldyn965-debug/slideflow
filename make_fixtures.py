"""Generates test fixtures with python-pptx (independent of the app's parser) plus deliberately broken files."""
import io, zipfile, struct, zlib
from pptx import Presentation
from pptx.util import Inches, Pt, Emu

def png(w=40, h=24, rgb=(120, 160, 200)):
    raw = b''.join(b'\x00' + bytes(rgb) * w for _ in range(h))
    def chunk(t, d): return struct.pack('>I', len(d)) + t + d + struct.pack('>I', zlib.crc32(t + d) & 0xffffffff)
    return b'\x89PNG\r\n\x1a\n' + chunk(b'IHDR', struct.pack('>IIBBBBB', w, h, 8, 2, 0, 0, 0)) + chunk(b'IDAT', zlib.compress(raw)) + chunk(b'IEND', b'')

out = 'fixtures/'

# 1. Widescreen deck: title slide, content w/ placeholder inheritance, picture, table, group, hidden slide
p = Presentation(); p.slide_width, p.slide_height = Inches(13.333), Inches(7.5)
s = p.slides.add_slide(p.slide_layouts[0]); s.shapes.title.text = 'Quarterly Review'; s.placeholders[1].text = 'Finance & Ops <b>2026</b>'
s = p.slides.add_slide(p.slide_layouts[1]); s.shapes.title.text = 'Highlights'; s.placeholders[1].text_frame.text = 'Revenue up 12%'
s.placeholders[1].text_frame.add_paragraph().text = 'Churn down'
s.shapes.add_picture(io.BytesIO(png()), Inches(8), Inches(3), Inches(3), Inches(1.8))
s = p.slides.add_slide(p.slide_layouts[5]); s.shapes.title.text = 'Numbers'
t = s.shapes.add_table(2, 2, Inches(1), Inches(2), Inches(6), Inches(1.5)).table
t.cell(0,0).text, t.cell(0,1).text, t.cell(1,0).text, t.cell(1,1).text = 'Q1','Q2','10','14'
g = s.shapes.add_group_shape(); g.shapes.add_textbox(Inches(8), Inches(2), Inches(4), Inches(1)).text_frame.text = 'In a group'
s = p.slides.add_slide(p.slide_layouts[6]); s.shapes.add_textbox(Inches(1), Inches(1), Inches(5), Inches(1)).text_frame.text = 'Hidden backup slide'
s._element.set('show', '0')
p.save(out + 'wide.pptx')

# 2. 4:3 deck, unicode text
p = Presentation()
s = p.slides.add_slide(p.slide_layouts[0]); s.shapes.title.text = 'مرحبا — Café ☕ 日本語'
p.save(out + 'standard43.pptx')

# 3. Linked picture (external) -> warning
p = Presentation(); s = p.slides.add_slide(p.slide_layouts[6])
s.shapes.add_picture(io.BytesIO(png()), Inches(1), Inches(1), Inches(2), Inches(1))
p.save('/tmp/ext.pptx')
zin = zipfile.ZipFile('/tmp/ext.pptx'); zout = zipfile.ZipFile(out + 'external-image.pptx', 'w', zipfile.ZIP_DEFLATED)
for n in zin.namelist():
    d = zin.read(n)
    if n.startswith('ppt/slides/_rels/'):
        d = d.replace(b'Target="../media/image1.png"', b'Target="https://example.com/x.png" TargetMode="External"')
    zout.writestr(n, d)
zout.close()

# --- broken files ---
open(out + 'not-a-pptx.pptx', 'w').write('hello, I am a text file')
open(out + 'empty.pptx', 'wb').close()
open(out + 'legacy-ole.pptx', 'wb').write(b'\xd0\xcf\x11\xe0\xa1\xb1\x1a\xe1' + b'\0' * 512)
open(out + 'wrong-extension.ppt', 'wb').write(open(out + 'wide.pptx', 'rb').read())
with zipfile.ZipFile(out + 'plain.zip.pptx', 'w') as z: z.writestr('hello.txt', 'hi')
zin = zipfile.ZipFile(out + 'wide.pptx'); zout = zipfile.ZipFile(out + 'macro.pptx', 'w', zipfile.ZIP_DEFLATED)
for n in zin.namelist():
    d = zin.read(n)
    if n == '[Content_Types].xml': d = d.replace(b'application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml', b'application/vnd.ms-powerpoint.presentation.macroEnabled.main+xml')
    zout.writestr(n, d)
zout.close()
zin = zipfile.ZipFile(out + 'wide.pptx'); zout = zipfile.ZipFile(out + 'broken-xml.pptx', 'w', zipfile.ZIP_DEFLATED)
for n in zin.namelist():
    d = zin.read(n)
    if n == 'ppt/slides/slide2.xml': d = d[:200]
    zout.writestr(n, d)
zout.close()
p = Presentation(); p.save(out + 'no-slides.pptx')
print('fixtures ok')
