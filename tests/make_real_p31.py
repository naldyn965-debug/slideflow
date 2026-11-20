"""Phase 3.1 'real-world' decks (tests/real/). These are NOT hand-built XML:
  ppt-default-*.pptx  python-pptx's bundled default.pptx: master, layouts and theme were authored in Microsoft PowerPoint (MIT-licensed with python-pptx);
                      only the slides on top are generated here. English and Arabic RTL variants.
  lo-*.pptx           the Phase 3.1 fixtures re-saved by LibreOffice Impress (soffice --convert-to pptx): every XML part is written by LibreOffice's
                      own exporter, which is a different writer from python-pptx (different ids, group handling, custom geometry, text properties).
Not available in the authoring sandbox (no network): files exported by PowerPoint itself, Keynote or Google Slides. Drop any such deck into tests/real/
and `node tests/render.mjs` includes it automatically (see tests/real/README.md).
Run from the repo root (needs python-pptx and, for the lo-* decks, soffice):  python3 tests/make_real_p31.py
"""
import subprocess, shutil, os, tempfile
from pptx import Presentation
from pptx.oxml import parse_xml

A = 'xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"'
OUT = 'tests/real/'


def english():
    p = Presentation()
    s = p.slides.add_slide(p.slide_layouts[0]); s.shapes.title.text = 'Quarterly Review'; s.placeholders[1].text = 'Prepared with the PowerPoint default master'
    s = p.slides.add_slide(p.slide_layouts[1]); s.shapes.title.text = 'Highlights'
    tf = s.placeholders[1].text_frame; tf.text = 'Revenue up 12% on the previous quarter'
    for t, l in (('Subscriptions led the growth', 1), ('Two new regions launched', 0), ('Support backlog cleared', 0)):
        q = tf.add_paragraph(); q.text = t; q.level = l
    s = p.slides.add_slide(p.slide_layouts[3]); s.shapes.title.text = 'Two content'
    s.placeholders[1].text_frame.text = 'Left column text'; s.placeholders[2].text_frame.text = 'Right column text'
    s = p.slides.add_slide(p.slide_layouts[5]); s.shapes.title.text = 'Title only'
    s = p.slides.add_slide(p.slide_layouts[2]); s.shapes.title.text = 'Section header'; s.placeholders[1].text = 'Section text'
    p.save(OUT + 'ppt-default-english.pptx')


def arabic():
    p = Presentation()
    def rtl(tf, texts, size=None):
        body = tf._txBody
        for pa in list(body.findall(f'{{http://schemas.openxmlformats.org/drawingml/2006/main}}p')): body.remove(pa)
        for t, lvl in texts:
            sz = f' sz="{size}"' if size else ''
            body.append(parse_xml(f'<a:p {A}><a:pPr algn="r" rtl="1" lvl="{lvl}"/><a:r><a:rPr lang="ar-SA"{sz}><a:latin typeface="Arial"/><a:cs typeface="Arial"/></a:rPr><a:t>{t}</a:t></a:r></a:p>'))
    s = p.slides.add_slide(p.slide_layouts[0]); rtl(s.shapes.title.text_frame, [('منصة نابتكس الزراعية', 0)]); rtl(s.placeholders[1].text_frame, [('مراجعة الربع الأخير', 0)])
    s = p.slides.add_slide(p.slide_layouts[1]); rtl(s.shapes.title.text_frame, [('أهم النتائج', 0)])
    rtl(s.placeholders[1].text_frame, [('ارتفع عدد الطلاب المسجلين بنسبة ١٢٪', 0), ('الدورات الجديدة في علوم النبات', 1), ('إطلاق نظام الدفع الجديد', 0), ('Nabtex Academy v2', 0)])
    s = p.slides.add_slide(p.slide_layouts[3]); rtl(s.shapes.title.text_frame, [('عمودان', 0)])
    rtl(s.placeholders[1].text_frame, [('العمود الأيمن', 0)]); rtl(s.placeholders[2].text_frame, [('العمود الأيسر', 0)])
    p.save(OUT + 'ppt-default-arabic-rtl.pptx')


def libreoffice():
    if not shutil.which('soffice'): print('soffice not found: skipping lo-* decks'); return
    tmp = tempfile.mkdtemp()
    for name in ('p31-text', 'p31-arabic', 'p31-shapes', 'p31-geometry', 'p31-template-decor', 'p31-media'):
        src = f'tests/fixtures/{name}.pptx'; d = os.path.join(tmp, name); os.makedirs(d)
        subprocess.run(['soffice', '--headless', '--convert-to', 'pptx:Impress MS PowerPoint 2007 XML', '--outdir', d, src], check=True, capture_output=True, timeout=180)
        shutil.copy(os.path.join(d, name + '.pptx'), OUT + 'lo-' + name[4:] + '.pptx')


if __name__ == '__main__':
    english(); arabic(); libreoffice()
    print(sorted(os.listdir(OUT)))
