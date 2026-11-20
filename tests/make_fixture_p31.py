"""Phase 3.1 fixtures: ground truth for the scene model and the Rendered tab (tests/render.mjs).
Run from the repo root:  python3 tests/make_fixture_p31.py   (needs python-pptx, Pillow and XlsxWriter for the chart).

  p31-geometry.pptx      nested groups with rotation / flips / non-uniform scale, connectors, z-order.
                         Expected world corners are computed HERE (independent of js/model/scene.js) and written to p31-geometry.expected.json
  p31-shapes.pptx        preset geometries (supported and not), custom geometry, gradients, dashes, arrowheads, shadow, text in shapes, flipV text
  p31-text.pptx          runs, bullets, spacing, anchors, insets, autofit, columns, vertical text, hyperlink, field, placeholders (master inheritance)
  p31-arabic.pptx        Arabic RTL (placeholders and text boxes, mixed Arabic/English, numbers, bullets, table, shape text) next to English
  p31-media.pptx         crop, rotation, flips, rounded picture, picture fill, unsupported media format, linked picture
  p31-unsupported.pptx   chart, table (merges, fills), SmartArt / OLE / content part (synthetic frames), mc:AlternateContent, unknown element,
                         duplicate and missing ids, a placeholder with no resolvable position
  p31-template-decor.pptx master + layout decoration, gradient master background, a slide that hides master shapes, a picture background
Every shape is raw XML so the expected values are exact. Synthetic frames (SmartArt, OLE, content part) point at relationships that do not exist:
they exist to prove that such content is listed and kept, not that PowerPoint can open them.
"""
import io, json, math, struct, zlib
from lxml import etree
from pptx import Presentation
from pptx.util import Emu
from pptx.oxml import parse_xml
from pptx.chart.data import CategoryChartData
from pptx.enum.chart import XL_CHART_TYPE

OUT = 'tests/fixtures/'
NSDECL = ('xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" '
          'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" '
          'xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" xmlns:p14="http://schemas.microsoft.com/office/powerpoint/2010/main"')
W, H = 12192000, 6858000


def X(xml):
    """Parse one fragment; the namespace declarations ride on the root element and python-pptx / lxml hoists them when it is appended."""
    i = xml.index('>') if xml[xml.index('<') + 1] != '/' else 0
    head, rest = xml[:i], xml[i:]
    if head.endswith('/'): head, rest = head[:-1], '/' + rest
    return parse_xml(f'{head} {NSDECL}{rest}')


def put(slide, *frags):
    tree = slide.shapes._spTree
    for f in frags: tree.append(X(f))


def png(rgb, w=40, h=24):
    raw = b''.join(b'\x00' + bytes(rgb) * w for _ in range(h))
    def chunk(t, d): return struct.pack('>I', len(d)) + t + d + struct.pack('>I', zlib.crc32(t + d) & 0xffffffff)
    return b'\x89PNG\r\n\x1a\n' + chunk(b'IHDR', struct.pack('>IIBBBBB', w, h, 8, 2, 0, 0, 0)) + chunk(b'IDAT', zlib.compress(raw)) + chunk(b'IEND', b'')


def gradient_png(w=200, h=120):
    """Left-to-right red->blue with a white corner marker, so crops and flips are visible."""
    rows = []
    for y in range(h):
        row = b'\x00'
        for x in range(w):
            r = int(255 * (1 - x / (w - 1))); b = int(255 * x / (w - 1))
            row += bytes((255, 255, 255)) if (x < 24 and y < 24) else bytes((r, 60, b))
        rows.append(row)
    raw = b''.join(rows)
    def chunk(t, d): return struct.pack('>I', len(d)) + t + d + struct.pack('>I', zlib.crc32(t + d) & 0xffffffff)
    return b'\x89PNG\r\n\x1a\n' + chunk(b'IHDR', struct.pack('>IIBBBBB', w, h, 8, 2, 0, 0, 0)) + chunk(b'IDAT', zlib.compress(raw)) + chunk(b'IEND', b'')


# ---------- raw XML builders ----------
def xfrm(x, y, w, h, rot=0, fh=False, fv=False, tag='a:xfrm', ch=None):
    a = (f' rot="{int(rot * 60000)}"' if rot else '') + (' flipH="1"' if fh else '') + (' flipV="1"' if fv else '')
    c = f'<a:chOff x="{ch[0]}" y="{ch[1]}"/><a:chExt cx="{ch[2]}" cy="{ch[3]}"/>' if ch else ''
    return f'<{tag}{a}><a:off x="{x}" y="{y}"/><a:ext cx="{w}" cy="{h}"/>{c}</{tag}>'


SOLID = lambda hx, extra='': f'<a:solidFill><a:srgbClr val="{hx}">{extra}</a:srgbClr></a:solidFill>'
LINE = lambda hx='1F3864', w=12700, extra='': f'<a:ln w="{w}">{SOLID(hx)}{extra}</a:ln>'


def run(t, sz=None, b=None, i=None, u=None, color=None, font=None, cs=None, lang='en-US', extra='', attrs=''):
    a = f' lang="{lang}"' + (f' sz="{int(sz * 100)}"' if sz else '') + (f' b="{int(b)}"' if b is not None else '') + (f' i="{int(i)}"' if i is not None else '') + (f' u="{u}"' if u else '') + attrs
    kids = (SOLID(color) if color else '') + (f'<a:latin typeface="{font}"/>' if font else '') + (f'<a:cs typeface="{cs}"/>' if cs else '') + extra
    t = t.replace('&', '&amp;').replace('<', '&lt;')
    return f'<a:r><a:rPr{a}>{kids}</a:rPr><a:t>{t}</a:t></a:r>'


def para(*runs, algn=None, rtl=None, lvl=None, marL=None, indent=None, bullet=None, spc=None, endsz=None):
    a = (f' algn="{algn}"' if algn else '') + (f' rtl="{int(rtl)}"' if rtl is not None else '') + (f' lvl="{lvl}"' if lvl else '') + (f' marL="{marL}"' if marL is not None else '') + (f' indent="{indent}"' if indent is not None else '')
    kids = (spc or '') + (bullet or '')
    return f'<a:p><a:pPr{a}>{kids}</a:pPr>{"".join(runs)}</a:p>'


def body(*paras, bodyPr='<a:bodyPr wrap="square" rtlCol="0"><a:noAutofit/></a:bodyPr>', tag='p:txBody'):
    return f'<{tag}>{bodyPr}<a:lstStyle/>{"".join(paras) or "<a:p><a:endParaRPr lang=\"en-US\"/></a:p>"}</{tag}>'


def sp(id, name, x, y, w, h, prst='rect', rot=0, fh=False, fv=False, fill=None, ln=None, tx=None, av='', geom=None, txbox=False, extra='', style=''):
    g = geom or f'<a:prstGeom prst="{prst}"><a:avLst>{av}</a:avLst></a:prstGeom>'
    f = SOLID('4472C4') if fill is None else fill
    l = LINE() if ln is None else ln
    return (f'<p:sp><p:nvSpPr><p:cNvPr id="{id}" name="{name}"/><p:cNvSpPr{" txBox=\"1\"" if txbox else ""}/><p:nvPr/></p:nvSpPr>'
            f'<p:spPr>{xfrm(x, y, w, h, rot, fh, fv)}{g}{f}{l}{extra}</p:spPr>{style}{tx if tx is not None else body()}</p:sp>')


def tbox(id, name, x, y, w, h, *paras, rot=0, bodyPr=None, fill='<a:noFill/>'):
    bp = bodyPr or '<a:bodyPr wrap="square" rtlCol="0"><a:noAutofit/></a:bodyPr>'
    return sp(id, name, x, y, w, h, fill=fill, ln='<a:ln><a:noFill/></a:ln>', tx=body(*paras, bodyPr=bp), txbox=True, rot=rot)


def grp(id, name, x, y, w, h, ch, *kids, rot=0, fh=False, fv=False):
    return (f'<p:grpSp><p:nvGrpSpPr><p:cNvPr id="{id}" name="{name}"/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>'
            f'<p:grpSpPr>{xfrm(x, y, w, h, rot, fh, fv, ch=ch)}</p:grpSpPr>{"".join(kids)}</p:grpSp>')


def cxn(id, name, x, y, w, h, prst='straightConnector1', rot=0, fh=False, fv=False, ln=None, av=''):
    return (f'<p:cxnSp><p:nvCxnSpPr><p:cNvPr id="{id}" name="{name}"/><p:cNvCxnSpPr/><p:nvPr/></p:nvCxnSpPr>'
            f'<p:spPr>{xfrm(x, y, w, h, rot, fh, fv)}<a:prstGeom prst="{prst}"><a:avLst>{av}</a:avLst></a:prstGeom>{ln or LINE("C00000", 28575)}</p:spPr></p:cxnSp>')


def new_deck(w=W, h=H):
    p = Presentation(); p.slide_width, p.slide_height = Emu(w), Emu(h); return p


def blank(p, layout=6): return p.slides.add_slide(p.slide_layouts[layout])


def label(id, x, y, text, sz=10, color='444444'):
    return tbox(id, f'label-{id}', x, y, 1800000, 260000, para(run(text, sz, color=color)))


# ======================================================================================================
# p31-geometry: independent expected world corners
# ======================================================================================================
def L(x, y, w, h, rot=0, fh=False, fv=False):
    """Local box (u, v in 0..w, 0..h) -> parent space: flip about the box centre, then rotate clockwise about it (y axis down)."""
    cx, cy, th = x + w / 2, y + h / 2, math.radians(rot)
    def f(u, v):
        dx, dy = u - w / 2, v - h / 2
        if fh: dx = -dx
        if fv: dy = -dy
        return (cx + dx * math.cos(th) - dy * math.sin(th), cy + dx * math.sin(th) + dy * math.cos(th))
    return f


def G(x, y, w, h, ch, rot=0, fh=False, fv=False):
    """Group: child-space point -> parent space. Scale chExt->ext about chOff->off, then flip/rotate the whole group about its centre."""
    cx0, cy0, cw, chh = ch
    gcx, gcy, th = x + w / 2, y + h / 2, math.radians(rot)
    def f(px, py):
        ux, uy = x + (px - cx0) * w / cw, y + (py - cy0) * h / chh
        dx, dy = ux - gcx, uy - gcy
        if fh: dx = -dx
        if fv: dy = -dy
        return (gcx + dx * math.cos(th) - dy * math.sin(th), gcy + dx * math.sin(th) + dy * math.cos(th))
    return f


def corners(box, local, *maps):
    x, y, w, h = box[:4]
    pts = [local(u, v) for u, v in ((0, 0), (w, 0), (w, h), (0, h))]
    for m in maps: pts = [m(*q) for q in pts]
    return [round(c, 3) for q in pts for c in q]


def make_geometry():
    p = new_deck(); s = blank(p)
    exp = {}
    z = [0]

    def rec(name, id, box, local, *maps, kind='shape'):
        exp[name] = {'id': id, 'corners': corners(box, local, *maps), 'z': z[0], 'kind': kind}; z[0] += 1

    # row 1: single shapes
    specs = [('g01-plain', 11, (500000, 400000, 1500000, 800000), {}), ('g02-rot30', 12, (2500000, 400000, 1500000, 800000), dict(rot=30)),
             ('g03-flipH', 13, (4500000, 400000, 1500000, 800000), dict(fh=True, prst='rightArrow')), ('g04-flipV', 14, (6500000, 400000, 1500000, 800000), dict(fv=True, prst='triangle')),
             ('g05-flipH-rot45', 15, (8500000, 400000, 1500000, 800000), dict(fh=True, rot=45, prst='rightArrow'))]
    for name, id, box, kw in specs:
        put(s, sp(id, name, *box, prst=kw.get('prst', 'rect'), rot=kw.get('rot', 0), fh=kw.get('fh', False), fv=kw.get('fv', False)))
        rec(name, id, box, L(*box, kw.get('rot', 0), kw.get('fh', False), kw.get('fv', False)))
    # g06: group, no rotation, scale 0.5
    g6 = dict(x=500000, y=1800000, w=3000000, h=2000000, ch=(0, 0, 6000000, 4000000))
    c61, c62 = (1000000, 1000000, 2000000, 1000000), (3500000, 2000000, 1500000, 1000000)
    put(s, grp(20, 'g06-group-scaled', g6['x'], g6['y'], g6['w'], g6['h'], g6['ch'],
               sp(21, 'g06-child1', *c61), sp(22, 'g06-child2-rot20', *c62, rot=20)))
    gm = G(**g6)
    rec('g06-group-scaled', 20, (g6['x'], g6['y'], g6['w'], g6['h']), L(g6['x'], g6['y'], g6['w'], g6['h']), kind='group')
    rec('g06-child1', 21, c61, L(*c61), gm); rec('g06-child2-rot20', 22, c62, L(*c62, 20), gm)
    # g07: group rotated 30 with an offset child space
    g7 = dict(x=4000000, y=1800000, w=2400000, h=1600000, ch=(1000000, 1000000, 2400000, 1600000), rot=30)
    c7 = (1200000, 1200000, 1000000, 600000)
    put(s, grp(30, 'g07-group-rot30', g7['x'], g7['y'], g7['w'], g7['h'], g7['ch'], sp(31, 'g07-child', *c7), rot=30))
    gm7 = G(**g7)
    rec('g07-group-rot30', 30, (g7['x'], g7['y'], g7['w'], g7['h']), L(g7['x'], g7['y'], g7['w'], g7['h'], 30), kind='group')
    rec('g07-child', 31, c7, L(*c7), gm7)
    # g08: nested groups, outer rot20+flipH with scale, inner rot40, child rot10
    o = dict(x=6900000, y=1800000, w=2400000, h=2000000, ch=(0, 0, 4800000, 4000000), rot=20, fh=True)
    i = dict(x=400000, y=400000, w=2400000, h=1600000, ch=(1000000, 1000000, 1200000, 800000), rot=40)
    c8 = (1200000, 1100000, 600000, 400000)
    inner = grp(41, 'g08-inner', i['x'], i['y'], i['w'], i['h'], i['ch'], sp(42, 'g08-child-rot10', *c8, rot=10), rot=40)
    put(s, grp(40, 'g08-outer', o['x'], o['y'], o['w'], o['h'], o['ch'], inner, rot=20, fh=True))
    go, gi = G(**o), G(**i)
    rec('g08-outer', 40, (o['x'], o['y'], o['w'], o['h']), L(o['x'], o['y'], o['w'], o['h'], 20, True), kind='group')
    rec('g08-inner', 41, (i['x'], i['y'], i['w'], i['h']), L(i['x'], i['y'], i['w'], i['h'], 40), go, kind='group')
    rec('g08-child-rot10', 42, c8, L(*c8, 10), gi, go)
    # g09: group flipped horizontally
    g9 = dict(x=9700000, y=1800000, w=1800000, h=1200000, ch=(0, 0, 1800000, 1200000), fh=True)
    c9 = (100000, 100000, 900000, 500000)
    put(s, grp(50, 'g09-group-flipH', g9['x'], g9['y'], g9['w'], g9['h'], g9['ch'], sp(51, 'g09-child', *c9, prst='rightArrow'), fh=True))
    rec('g09-group-flipH', 50, (g9['x'], g9['y'], g9['w'], g9['h']), L(g9['x'], g9['y'], g9['w'], g9['h'], 0, True), kind='group')
    rec('g09-child', 51, c9, L(*c9), G(**g9))
    # connectors and lines
    k1 = (500000, 4700000, 2500000, 1200000); put(s, cxn(60, 'g10-connector-flipV', *k1, fv=True)); rec('g10-connector-flipV', 60, k1, L(*k1, 0, False, True), kind='line')
    k2 = (3500000, 4700000, 2500000, 1200000); put(s, cxn(61, 'g11-bent-flipH', *k2, prst='bentConnector3', fh=True, av='<a:gd name="adj1" fmla="val 30000"/>')); rec('g11-bent-flipH', 61, k2, L(*k2, 0, True), kind='line')
    k3 = (7000000, 4700000, 1500000, 0); put(s, cxn(62, 'g12-line-rot90', *k3, prst='line', rot=90)); rec('g12-line-rot90', 62, k3, L(*k3, 90), kind='line')
    # text box rotated -10 degrees, and a later shape that overlaps g01 to pin z-order
    t = (9000000, 4800000, 2400000, 700000)
    put(s, tbox(63, 'g13-textbox-rot350', *t, para(run('Rotated -10', 18)), rot=350)); rec('g13-textbox-rot350', 63, t, L(*t, 350), kind='text')
    ov = (800000, 600000, 700000, 400000)
    put(s, sp(64, 'g14-overlap-on-top', *ov, fill=SOLID('ED7D31', '<a:alpha val="60000"/>'))); rec('g14-overlap-on-top', 64, ov, L(*ov))
    p.save(OUT + 'p31-geometry.pptx'); print('wrote p31-geometry.pptx')
    json.dump({'slide': {'width': W, 'height': H}, 'elements': exp}, open(OUT + 'p31-geometry.expected.json', 'w'), indent=1)
    print('wrote p31-geometry.expected.json')


# ======================================================================================================
def make_shapes():
    p = new_deck()
    s = blank(p)
    presets = ['rect', 'roundRect', 'ellipse', 'triangle', 'rtTriangle', 'diamond', 'parallelogram', 'trapezoid', 'hexagon', 'homePlate', 'chevron',
               'rightArrow', 'leftArrow', 'upArrow', 'downArrow', 'plus', 'octagon', 'flowChartProcess', 'flowChartDecision', 'flowChartTerminator',
               'star5', 'cloud', 'heart', 'sun', 'lightningBolt', 'wedgeRectCallout', 'donut', 'blockArc']
    cols, cw, ch, gx, gy = 7, 1450000, 900000, 250000, 450000
    for k, pr in enumerate(presets):
        x = 350000 + (k % cols) * (cw + gx); y = 300000 + (k // cols) * (ch + gy)
        put(s, sp(100 + k, f'preset-{pr}', x, y, cw, ch, prst=pr, fill=SOLID('2E75B6'), tx=body(para(run(pr, 9, color='FFFFFF'), algn='ctr'), bodyPr='<a:bodyPr anchor="ctr"/>')))
    s2 = blank(p)
    # custom geometry: moveTo/lnTo/cubicBez/arcTo/close in a 1000x1000 path space; and one that uses a guide name (-> approximated)
    cust = ('<a:custGeom><a:avLst/><a:gdLst/><a:ahLst/><a:cxnLst/><a:rect l="0" t="0" r="r" b="b"/><a:pathLst><a:path w="1000" h="1000">'
            '<a:moveTo><a:pt x="0" y="1000"/></a:moveTo><a:lnTo><a:pt x="500" y="0"/></a:lnTo><a:cubicBezTo><a:pt x="800" y="300"/><a:pt x="1000" y="500"/><a:pt x="1000" y="1000"/></a:cubicBezTo><a:close/></a:path>'
            '<a:path w="1000" h="1000" fill="none"><a:moveTo><a:pt x="100" y="800"/></a:moveTo><a:arcTo wR="300" hR="300" stAng="10800000" swAng="10800000"/></a:path></a:pathLst></a:custGeom>')
    put(s2, sp(200, 'cust-path', 400000, 300000, 1800000, 1500000, geom=cust, fill=SOLID('70AD47')))
    guide = cust.replace('<a:gdLst/>', '<a:gdLst><a:gd name="g1" fmla="*/ 1000 1 2"/><a:gd name="g2" fmla="+- g1 0 0"/></a:gdLst>').replace('<a:pt x="500" y="0"/>', '<a:pt x="g2" y="0"/>')
    put(s2, sp(201, 'cust-guide', 2600000, 300000, 1800000, 1500000, geom=guide, fill=SOLID('A5A5A5')))
    grad = '<a:gradFill rotWithShape="1"><a:gsLst><a:gs pos="0"><a:srgbClr val="FF0000"/></a:gs><a:gs pos="100000"><a:srgbClr val="0000FF"/></a:gs></a:gsLst><a:lin ang="2700000" scaled="0"/></a:gradFill>'
    put(s2, sp(202, 'fill-linear-gradient', 4800000, 300000, 1800000, 1500000, fill=grad))
    rad = '<a:gradFill><a:gsLst><a:gs pos="0"><a:srgbClr val="FFFF00"/></a:gs><a:gs pos="100000"><a:srgbClr val="C00000"/></a:gs></a:gsLst><a:path path="circle"><a:fillToRect l="50000" t="50000" r="50000" b="50000"/></a:path></a:gradFill>'
    put(s2, sp(203, 'fill-radial-gradient', 7000000, 300000, 1800000, 1500000, fill=rad))
    patt = '<a:pattFill prst="ltUpDiag"><a:fgClr><a:srgbClr val="000000"/></a:fgClr><a:bgClr><a:srgbClr val="FFFFFF"/></a:bgClr></a:pattFill>'
    put(s2, sp(204, 'fill-pattern', 9200000, 300000, 1800000, 1500000, fill=patt))
    for k, (dash, w) in enumerate([('solid', 38100), ('dash', 38100), ('sysDot', 38100), ('lgDashDot', 38100)]):
        put(s2, sp(210 + k, f'line-{dash}', 400000 + k * 2800000, 2300000, 2300000, 600000, fill='<a:noFill/>', ln=f'<a:ln w="{w}" cap="rnd">{SOLID("1F3864")}<a:prstDash val="{dash}"/><a:round/></a:ln>'))
    put(s2, cxn(220, 'arrow-ends', 400000, 3300000, 3000000, 0, prst='straightConnector1', ln=f'<a:ln w="38100">{SOLID("C00000")}<a:headEnd type="triangle"/><a:tailEnd type="arrow" w="lg" len="lg"/></a:ln>'))
    put(s2, sp(221, 'shadow-outer', 4200000, 3000000, 2000000, 900000, extra='<a:effectLst><a:outerShdw blurRad="76200" dist="76200" dir="2700000" algn="tl" rotWithShape="0"><a:srgbClr val="000000"><a:alpha val="50000"/></a:srgbClr></a:outerShdw></a:effectLst>'))
    put(s2, sp(222, 'glow-effect', 6800000, 3000000, 2000000, 900000, extra='<a:effectLst><a:glow rad="101600"><a:srgbClr val="FFC000"><a:alpha val="60000"/></a:srgbClr></a:glow></a:effectLst>'))
    put(s2, sp(223, 'text-flipV', 9200000, 3000000, 2200000, 900000, fv=True, tx=body(para(run('Flipped', 18, color='FFFFFF'), algn='ctr'), bodyPr='<a:bodyPr anchor="ctr"/>')))
    put(s2, sp(224, 'style-ref-fill', 400000, 4300000, 2200000, 900000, fill='', ln='',
               style='<p:style><a:lnRef idx="2"><a:schemeClr val="accent1"><a:shade val="50000"/></a:schemeClr></a:lnRef><a:fillRef idx="1"><a:schemeClr val="accent2"/></a:fillRef><a:effectRef idx="0"><a:schemeClr val="accent1"/></a:effectRef><a:fontRef idx="minor"><a:schemeClr val="lt1"/></a:fontRef></p:style>',
               tx=body(para(run('Theme style', 16), algn='ctr'), bodyPr='<a:bodyPr anchor="ctr"/>')))
    put(s2, sp(225, 'rot-roundRect-adj', 3100000, 4300000, 2200000, 900000, prst='roundRect', rot=15, av='<a:gd name="adj" fmla="val 50000"/>', fill=SOLID('7030A0')))
    p.save(OUT + 'p31-shapes.pptx'); print('wrote p31-shapes.pptx')


# ======================================================================================================
def make_text():
    p = new_deck()
    s = blank(p)
    bu = lambda ch='•': f'<a:buFont typeface="Arial"/><a:buChar char="{ch}"/>'
    put(s,
        tbox(300, 'mixed-runs', 400000, 300000, 5400000, 1200000,
             para(run('Bold ', 24, b=1), run('italic ', 24, i=1), run('underline ', 24, u='sng'), run('red ', 24, color='C00000'), run('Georgia', 24, font='Georgia'), run(' x', 24), run('2', 24, attrs=' baseline="30000"'), run(' H', 24), run('2', 24, attrs=' baseline="-25000"'), run('O', 24)),
             para(run('Second paragraph, centred, 14 pt, strike', 14, attrs=' strike="sngStrike"'), algn='ctr')),
        tbox(301, 'bullets', 6200000, 300000, 5400000, 2200000,
             para(run('Level 1 bullet', 18), bullet=bu(), marL=342900, indent=-342900),
             para(run('Level 2 bullet', 16), lvl=1, bullet=bu('–'), marL=742950, indent=-285750),
             para(run('Numbered one', 18), bullet='<a:buAutoNum type="arabicPeriod"/>', marL=457200, indent=-457200),
             para(run('Numbered two', 18), bullet='<a:buAutoNum type="arabicPeriod"/>', marL=457200, indent=-457200),
             para(run('No bullet, long wrapped line to prove the hanging indent keeps its left edge when the text wraps onto a second line.', 14), marL=342900, indent=-342900, bullet='<a:buNone/>')),
        tbox(302, 'spacing', 400000, 1800000, 5400000, 1700000,
             para(run('Line spacing 150%', 16), spc='<a:lnSpc><a:spcPct val="150000"/></a:lnSpc><a:spcBef><a:spcPts val="600"/></a:spcBef>'),
             para(run('Line spacing exactly 30pt', 16), spc='<a:lnSpc><a:spcPts val="3000"/></a:lnSpc><a:spcAft><a:spcPts val="1200"/></a:spcAft>'),
             para(run('Justified paragraph with enough words to wrap onto several lines so justification is visible in the rendering of this slide.', 14), algn='just')),
        tbox(303, 'hyperlink-and-field', 6200000, 2700000, 5400000, 900000,
             f'<a:p><a:r><a:rPr lang="en-US" sz="1800"><a:hlinkClick r:id=""/></a:rPr><a:t>A hyperlink run</a:t></a:r><a:r><a:rPr lang="en-US" sz="1800"/><a:t> and field: </a:t></a:r><a:fld id="{{B6F15528-21DE-4FAA-801E-634DDDAF4B2B}}" type="slidenum"><a:rPr lang="en-US" sz="1800"/><a:t>7</a:t></a:fld></a:p>'),
        tbox(304, 'right-aligned', 400000, 3800000, 5400000, 500000, para(run('Right aligned', 18), algn='r')),
        tbox(305, 'highlight-and-caps', 6200000, 3800000, 5400000, 500000, para(run('Small caps text', 18, attrs=' cap="small" spc="200"'), run(' highlighted', 18, extra='<a:highlight><a:srgbClr val="FFFF00"/></a:highlight>'))),
        )
    s2 = blank(p)
    fr = lambda a, b: '<a:ln w="9525"><a:solidFill><a:srgbClr val="999999"/></a:solidFill></a:ln>'
    for k, (anc, name) in enumerate([('t', 'anchor-top'), ('ctr', 'anchor-middle'), ('b', 'anchor-bottom')]):
        put(s2, sp(310 + k, name, 400000 + k * 2000000, 300000, 1800000, 1200000, fill='<a:noFill/>', ln=fr(0, 0), txbox=True, tx=body(para(run(name, 14)), bodyPr=f'<a:bodyPr anchor="{anc}"><a:noAutofit/></a:bodyPr>')))
    put(s2,
        sp(320, 'insets-large', 6600000, 300000, 2300000, 1200000, fill='<a:noFill/>', ln=fr(0, 0), tx=body(para(run('Insets 0.3in / 0.1in', 14)), bodyPr='<a:bodyPr lIns="274320" tIns="91440" rIns="274320" bIns="91440"><a:noAutofit/></a:bodyPr>')),
        sp(321, 'norm-autofit-scale', 9100000, 300000, 2300000, 1200000, fill='<a:noFill/>', ln=fr(0, 0), tx=body(para(run('normAutofit fontScale 62.5% lnSpcReduction 20%: this text was shrunk by PowerPoint', 24)), bodyPr='<a:bodyPr><a:normAutofit fontScale="62500" lnSpcReduction="20000"/></a:bodyPr>')),
        sp(322, 'norm-autofit-noscale', 400000, 1800000, 2300000, 1200000, fill='<a:noFill/>', ln=fr(0, 0), tx=body(para(run('normAutofit without a stored scale: PowerPoint recomputes it', 24)), bodyPr='<a:bodyPr><a:normAutofit/></a:bodyPr>')),
        sp(323, 'shape-autofit', 2900000, 1800000, 2300000, 1200000, fill='<a:noFill/>', ln=fr(0, 0), tx=body(para(run('spAutoFit: the frame follows the text', 14)), bodyPr='<a:bodyPr><a:spAutoFit/></a:bodyPr>')),
        sp(324, 'columns-2', 5400000, 1800000, 3300000, 1200000, fill='<a:noFill/>', ln=fr(0, 0), tx=body(para(run('Two columns of text. The first column fills and then the text flows into the second column as it does in PowerPoint when numCol is two.', 12)), bodyPr='<a:bodyPr numCol="2" spcCol="182880"><a:noAutofit/></a:bodyPr>')),
        sp(325, 'vertical-vert', 8900000, 1800000, 700000, 2200000, fill='<a:noFill/>', ln=fr(0, 0), tx=body(para(run('Vertical', 16)), bodyPr='<a:bodyPr vert="vert"><a:noAutofit/></a:bodyPr>')),
        sp(326, 'vertical-vert270', 9800000, 1800000, 700000, 2200000, fill='<a:noFill/>', ln=fr(0, 0), tx=body(para(run('Vertical 270', 16)), bodyPr='<a:bodyPr vert="vert270"><a:noAutofit/></a:bodyPr>')),
        sp(327, 'no-wrap', 400000, 3300000, 1500000, 500000, fill='<a:noFill/>', ln=fr(0, 0), tx=body(para(run('wrap="none" text runs past its frame', 14)), bodyPr='<a:bodyPr wrap="none"><a:noAutofit/></a:bodyPr>')),
        sp(328, 'rotated-text-frame', 6600000, 3600000, 2800000, 800000, rot=-8 + 360, fill=SOLID('FFF2CC'), tx=body(para(run('Rotated frame', 18, color='000000'), algn='ctr'), bodyPr='<a:bodyPr anchor="ctr"/>')),
        )
    # placeholders: title + body inherit position, size, colour, bullets from the default master
    s3 = p.slides.add_slide(p.slide_layouts[1])
    s3.shapes.title.text_frame.text = 'Title from the master'
    s3.placeholders[1].text_frame.text = 'First level'
    q = s3.placeholders[1].text_frame.add_paragraph(); q.text = 'Second level'; q.level = 1
    p.save(OUT + 'p31-text.pptx'); print('wrote p31-text.pptx')


# ======================================================================================================
def make_arabic():
    p = new_deck()
    AR = lambda t, sz=24, **k: run(t, sz, lang='ar-SA', font=k.pop('font', 'Arial'), cs=k.pop('cs', 'Arial'), **k)
    s = p.slides.add_slide(p.slide_layouts[1])
    t = s.shapes.title.text_frame; t.text = ''
    t._txBody.remove(t._txBody.findall('{http://schemas.openxmlformats.org/drawingml/2006/main}p')[0])
    t._txBody.append(X(para(AR('عنوان العرض التقديمي', 40, b=1), algn='r', rtl=1)))
    b = s.placeholders[1].text_frame
    for pa in list(b._txBody.findall('{http://schemas.openxmlformats.org/drawingml/2006/main}p')): b._txBody.remove(pa)
    bu = '<a:buFont typeface="Arial"/><a:buChar char="•"/>'
    for frag in (para(AR('النقطة الأولى: الزراعة المستدامة', 28), algn='r', rtl=1, marL=342900, indent=-342900, bullet=bu),
                 para(AR('النقطة الثانية: الري بالتنقيط وخفض استهلاك المياه', 28), algn='r', rtl=1, marL=342900, indent=-342900, bullet=bu),
                 para(AR('مستوى فرعي', 24), algn='r', rtl=1, lvl=1, marL=742950, indent=-285750, bullet=bu)):
        b._txBody.append(X(frag))
    s2 = blank(p)
    put(s2,
        tbox(400, 'ar-rtl-right', 6200000, 300000, 5400000, 900000, para(AR('هذا نص عربي من اليمين إلى اليسار', 24), algn='r', rtl=1)),
        tbox(401, 'ar-wrapped', 6200000, 1300000, 5400000, 1400000, para(AR('فقرة عربية طويلة تلتف على عدة أسطر لإظهار اتجاه الكتابة والمحاذاة إلى اليمين بشكل صحيح داخل الإطار', 20), algn='r', rtl=1)),
        tbox(402, 'ar-mixed', 400000, 300000, 5400000, 900000, para(AR('الإصدار ', 22), run('Nabtex 2.5', 22, font='Calibri'), AR(' متاح الآن في 2026', 22), algn='r', rtl=1)),
        tbox(403, 'ar-numbers', 400000, 1300000, 5400000, 700000, para(AR('الأرقام: 12345 ٠١٢٣٤٥ (٪٥٠)', 22), algn='r', rtl=1)),
        tbox(404, 'en-left', 400000, 2100000, 5400000, 700000, para(run('English left-to-right text beside Arabic', 22))),
        tbox(405, 'ar-justified', 400000, 2900000, 5400000, 1300000, para(AR('فقرة عربية مضبوطة الأطراف تحتوي على كلمات كثيرة بما يكفي لإظهار الضبط على عدة أسطر', 18), algn='just', rtl=1)),
        tbox(406, 'ar-rtl-no-algn', 6200000, 2900000, 5400000, 700000, para(AR('محاذاة غير محددة', 22), rtl=1)),
        sp(407, 'ar-shape', 6200000, 3800000, 3000000, 900000, prst='roundRect', fill=SOLID('548235'), tx=body(para(AR('نص داخل شكل', 22, color='FFFFFF'), algn='ctr', rtl=1), bodyPr='<a:bodyPr anchor="ctr" rtlCol="1"/>')),
        )
    tbl = s2.shapes.add_table(2, 2, Emu(400000), Emu(4500000), Emu(5400000), Emu(1200000))
    for r, row in enumerate([('الصنف', 'العدد'), ('طماطم', '١٢')]):
        for c, txt in enumerate(row):
            cell = tbl.table.cell(r, c); tb = cell._tc.txBody
            for pa in list(tb.findall('{http://schemas.openxmlformats.org/drawingml/2006/main}p')): tb.remove(pa)
            tb.append(X(para(AR(txt, 18), algn='r', rtl=1)))
    p.save(OUT + 'p31-arabic.pptx'); print('wrote p31-arabic.pptx')


# ======================================================================================================
def make_media():
    import tempfile, os
    p = new_deck(); s = blank(p)
    d = tempfile.mkdtemp(); fp = os.path.join(d, 'g.png'); open(fp, 'wb').write(gradient_png())
    pics = []
    def add_pic(name, x, y, w, h, **kw):
        pc = s.shapes.add_picture(fp, Emu(x), Emu(y), Emu(w), Emu(h)); pc.name = name
        sppr = pc._element.spPr
        xf = sppr.find('{http://schemas.openxmlformats.org/drawingml/2006/main}xfrm')
        if kw.get('rot'): xf.set('rot', str(int(kw['rot'] * 60000)))
        if kw.get('fh'): xf.set('flipH', '1')
        if kw.get('fv'): xf.set('flipV', '1')
        bf = pc._element.blipFill
        if kw.get('crop'):
            l, t, r, b = kw['crop']; sr = X(f'<a:srcRect l="{l}" t="{t}" r="{r}" b="{b}"/>')
            bf.insert(1, sr)
        if kw.get('prst'): sppr.find('{http://schemas.openxmlformats.org/drawingml/2006/main}prstGeom').set('prst', kw['prst'])
        return pc
    add_pic('pic-plain', 400000, 300000, 2400000, 1440000)
    add_pic('pic-crop', 3200000, 300000, 2400000, 1440000, crop=(25000, 0, 25000, 30000))
    add_pic('pic-rot', 6000000, 300000, 2400000, 1440000, rot=15)
    add_pic('pic-flipH', 8900000, 300000, 2400000, 1440000, fh=True)
    add_pic('pic-ellipse', 400000, 2300000, 1800000, 1800000, prst='ellipse')
    add_pic('pic-flipV-rot', 3200000, 2300000, 2400000, 1440000, fv=True, rot=20)
    # picture fill inside an autoshape (blipFill on spPr)
    rid = s.part.get_or_add_image_part(fp)[1]
    put(s, sp(500, 'shape-picture-fill', 6000000, 2300000, 2400000, 1440000, fill=f'<a:blipFill><a:blip r:embed="{rid}"/><a:stretch><a:fillRect/></a:stretch></a:blipFill>'))
    # an image part whose format the browser cannot show (EMF magic bytes only) and a linked picture
    emf_part = s.part.package  # noqa
    from pptx.opc.package import Part
    from pptx.opc.packuri import PackURI
    emf = Part(PackURI('/ppt/media/image-fake.emf'), 'image/x-emf', s.part.package, b'\x01\x00\x00\x00' + b'\x00' * 80)
    rid2 = s.part.relate_to(emf, 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/image')
    put(s, f'<p:pic><p:nvPicPr><p:cNvPr id="501" name="pic-emf-unsupported"/><p:cNvPicPr/><p:nvPr/></p:nvPicPr><p:blipFill><a:blip r:embed="{rid2}"/><a:stretch><a:fillRect/></a:stretch></p:blipFill><p:spPr>{xfrm(8900000, 2300000, 2400000, 1440000)}<a:prstGeom prst="rect"><a:avLst/></a:prstGeom></p:spPr></p:pic>')
    rid3 = s.part.relate_to('https://example.invalid/linked.png', 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/image', is_external=True)
    put(s, f'<p:pic><p:nvPicPr><p:cNvPr id="502" name="pic-linked-external"/><p:cNvPicPr/><p:nvPr/></p:nvPicPr><p:blipFill><a:blip r:link="{rid3}"/><a:stretch><a:fillRect/></a:stretch></p:blipFill><p:spPr>{xfrm(400000, 4500000, 2400000, 1440000)}<a:prstGeom prst="rect"><a:avLst/></a:prstGeom></p:spPr></p:pic>')
    # duotone recolour (not rendered -> approximated)
    pr = add_pic('pic-duotone', 3200000, 4500000, 2400000, 1440000)
    pr._element.blipFill.find('{http://schemas.openxmlformats.org/drawingml/2006/main}blip').append(X('<a:duotone><a:srgbClr val="000000"/><a:srgbClr val="FFFFFF"/></a:duotone>'))
    p.save(OUT + 'p31-media.pptx'); print('wrote p31-media.pptx')


# ======================================================================================================
def make_unsupported():
    p = new_deck(); s = blank(p)
    cd = CategoryChartData(); cd.categories = ['A', 'B', 'C']; cd.add_series('S1', (1, 3, 2))
    gf = s.shapes.add_chart(XL_CHART_TYPE.COLUMN_CLUSTERED, Emu(400000), Emu(300000), Emu(4200000), Emu(2600000), cd); gf.name = 'chart-real'
    tb = s.shapes.add_table(3, 3, Emu(5000000), Emu(300000), Emu(3600000), Emu(1500000)); tb.name = 'table-merged'
    tbl = tb.table
    tbl.cell(0, 0).merge(tbl.cell(0, 1))
    tbl.cell(0, 0).text = 'Merged header'; tbl.cell(1, 0).text = 'a'; tbl.cell(1, 1).text = 'b'; tbl.cell(2, 2).text = 'c'
    tc = tbl.cell(1, 1)._tc
    tcPr = tc.get_or_add_tcPr(); tcPr.append(X(SOLID('FFE699')))
    # synthetic frames: SmartArt, OLE, content part (dangling relationships on purpose)
    put(s,
        f'<p:graphicFrame><p:nvGraphicFramePr><p:cNvPr id="600" name="smartart-synthetic"/><p:cNvGraphicFramePr/><p:nvPr/></p:nvGraphicFramePr>{xfrm(9000000, 300000, 2700000, 1800000, tag="p:xfrm")}'
        '<a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/diagram"><dgm:relIds xmlns:dgm="http://schemas.openxmlformats.org/drawingml/2006/diagram" r:dm="rId90" r:lo="rId91" r:qs="rId92" r:cs="rId93"/></a:graphicData></a:graphic></p:graphicFrame>',
        f'<p:graphicFrame><p:nvGraphicFramePr><p:cNvPr id="601" name="ole-synthetic"/><p:cNvGraphicFramePr/><p:nvPr/></p:nvGraphicFramePr>{xfrm(9000000, 2300000, 2700000, 1500000, tag="p:xfrm")}'
        '<a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/presentationml/2006/ole"><p:oleObj name="Worksheet" r:id="rId94" imgW="100" imgH="100" progId="Excel.Sheet.12"/></a:graphicData></a:graphic></p:graphicFrame>',
        '<p:contentPart r:id="rId95"/>',
        # AlternateContent: Choice and Fallback share an id on purpose
        '<mc:AlternateContent><mc:Choice Requires="p14">' + sp(610, 'alt-choice-shape', 400000, 3300000, 1800000, 900000, prst='roundRect', fill=SOLID('C55A11')) + '</mc:Choice>'
        '<mc:Fallback>' + sp(610, 'alt-fallback-shape', 400000, 3300000, 1800000, 900000, prst='ellipse', fill=SOLID('7F7F7F')) + '</mc:Fallback></mc:AlternateContent>',
        # AlternateContent whose Choice needs a graphic frame we cannot draw: the Fallback shape is drawn instead
        '<mc:AlternateContent><mc:Choice Requires="cx1"><p:graphicFrame><p:nvGraphicFramePr><p:cNvPr id="611" name="alt-chartex-choice"/><p:cNvGraphicFramePr/><p:nvPr/></p:nvGraphicFramePr>' + xfrm(6800000, 3300000, 1800000, 900000, tag="p:xfrm") + '<a:graphic><a:graphicData uri="http://schemas.microsoft.com/office/drawing/2014/chartex"/></a:graphic></p:graphicFrame></mc:Choice>'
        '<mc:Fallback>' + sp(611, 'alt-chartex-fallback', 6800000, 3300000, 1800000, 900000, prst='rect', fill=SOLID('A9D18E')) + '</mc:Fallback></mc:AlternateContent>',
        # unknown element with a position-less body and one unknown element with no children at all
        '<p:unknownVendorThing><p:cNvPr id="620" name="unknown-element"/></p:unknownVendorThing>',
        # duplicate ids and a missing id
        sp(630, 'dup-id-a', 2600000, 3300000, 1200000, 700000, fill=SOLID('2E75B6')), sp(630, 'dup-id-b', 4000000, 3300000, 1200000, 700000, fill=SOLID('00B0F0')),
        '<p:sp><p:nvSpPr><p:cNvPr name="no-id"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr><p:spPr>' + xfrm(5400000, 3300000, 1200000, 700000) + '<a:prstGeom prst="rect"><a:avLst/></a:prstGeom></p:spPr></p:sp>',
        # a placeholder that exists on no layout and has no xfrm: its position cannot be resolved
        '<p:sp><p:nvSpPr><p:cNvPr id="640" name="orphan-placeholder"/><p:cNvSpPr/><p:nvPr><p:ph type="hdr" idx="77"/></p:nvPr></p:nvSpPr><p:spPr/>' + body(para(run('Orphan placeholder text', 18))) + '</p:sp>',
        # a body placeholder whose idx the layout does not have: PowerPoint falls back to the master's body placeholder, matched by type
        '<p:sp><p:nvSpPr><p:cNvPr id="641" name="orphan-body-placeholder"/><p:cNvSpPr/><p:nvPr><p:ph type="body" idx="78"/></p:nvPr></p:nvSpPr><p:spPr/>' + body(para(run('Orphan body text', 18))) + '</p:sp>',
        # a picture-less blip shape with a bogus relationship id
        f'<p:pic><p:nvPicPr><p:cNvPr id="650" name="pic-bad-rel"/><p:cNvPicPr/><p:nvPr/></p:nvPicPr><p:blipFill><a:blip r:embed="rId999"/><a:stretch><a:fillRect/></a:stretch></p:blipFill><p:spPr>{xfrm(6800000, 3300000, 1200000, 700000)}<a:prstGeom prst="rect"><a:avLst/></a:prstGeom></p:spPr></p:pic>',
        )
    p.save(OUT + 'p31-unsupported.pptx'); print('wrote p31-unsupported.pptx')


# ======================================================================================================
def make_template_decor():
    p = new_deck(); s1 = p.slides.add_slide(p.slide_layouts[1]); s1.shapes.title.text_frame.text = 'Decor slide'; s1.placeholders[1].text_frame.text = 'Body text'
    s2 = p.slides.add_slide(p.slide_layouts[1]); s2.shapes.title.text_frame.text = 'Master shapes hidden'; s2.placeholders[1].text_frame.text = 'showMasterSp=0 on this slide'
    s2._element.set('showMasterSp', '0')
    s3 = p.slides.add_slide(p.slide_layouts[6]); s3.shapes._spTree  # blank layout, own picture background
    master = p.slide_master; layout = p.slide_layouts[1]
    mtree = master.shapes._spTree; ltree = layout.shapes._spTree
    mtree.insert(2, X(sp(900, 'master-bar', 0, 0, W, 300000, fill=SOLID('1F3864'), ln='<a:ln><a:noFill/></a:ln>')))
    ltree.insert(2, X(sp(901, 'layout-badge', W - 1800000, H - 900000, 1400000, 500000, prst='roundRect', fill=SOLID('ED7D31'), ln='<a:ln><a:noFill/></a:ln>',
                         tx=body(para(run('layout badge', 12, color='FFFFFF'), algn='ctr'), bodyPr='<a:bodyPr anchor="ctr"/>'))))
    cs = master._element.find('{http://schemas.openxmlformats.org/presentationml/2006/main}cSld')
    bg = X('<p:bg><p:bgPr><a:gradFill><a:gsLst><a:gs pos="0"><a:srgbClr val="F2F2F2"/></a:gs><a:gs pos="100000"><a:srgbClr val="BDD7EE"/></a:gs></a:gsLst><a:lin ang="5400000" scaled="0"/></a:gradFill><a:effectLst/></p:bgPr></p:bg>')
    old = cs.find('{http://schemas.openxmlformats.org/presentationml/2006/main}bg')
    if old is not None: cs.remove(old)
    cs.insert(0, bg)
    import tempfile, os
    d = tempfile.mkdtemp(); fp = os.path.join(d, 'g.png'); open(fp, 'wb').write(gradient_png())
    rid = s3.part.get_or_add_image_part(fp)[1]
    csl = s3._element.find('{http://schemas.openxmlformats.org/presentationml/2006/main}cSld')
    csl.insert(0, X(f'<p:bg><p:bgPr><a:blipFill><a:blip r:embed="{rid}"/><a:stretch><a:fillRect/></a:stretch></a:blipFill><a:effectLst/></p:bgPr></p:bg>'))
    put(s3, tbox(910, 'on-picture-bg', 1000000, 2500000, 6000000, 800000, para(run('Text over a picture background', 28, b=1, color='FFFFFF'))))
    p.save(OUT + 'p31-template-decor.pptx'); print('wrote p31-template-decor.pptx')


if __name__ == '__main__':
    make_geometry(); make_shapes(); make_text(); make_arabic(); make_media(); make_unsupported(); make_template_decor()
