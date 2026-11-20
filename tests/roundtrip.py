#!/usr/bin/env python3
"""Round-trip check (Phase 3.1). tests/render.mjs builds the scene model and renders every slide, then re-serialises the live OOXML documents of every slide,
layout and master it touched into a new .pptx (all other parts are copied untouched). This script proves that nothing was lost or changed:

  1. the package has exactly the same parts; every part that was not re-serialised (media, charts, SmartArt, OLE, embeddings, themes...) is byte-identical;
  2. every re-serialised XML part is canonically (C14N) equal to the original: same elements, attributes, namespaces, text, order;
  3. python-pptx opens both files and sees the same slides, shape ids, names, positions, rotations and text;
  4. LibreOffice renders both files identically (mean pixel error ~0).

usage: roundtrip.py OUT_JSON ORIGINAL.pptx ROUNDTRIP.pptx [ORIGINAL ROUNDTRIP ...]
"""
import json, os, shutil, subprocess, sys, tempfile, zipfile
from lxml import etree
from PIL import Image
import numpy as np

W = 640


def c14n(data):
    return etree.tostring(etree.fromstring(data), method='c14n')


def pptx_summary(path):
    from pptx import Presentation
    out = []
    prs = Presentation(path)
    for s in prs.slides:
        row = []
        def walk(shapes):
            for sh in shapes:
                try:
                    pos = (sh.left, sh.top, sh.width, sh.height, getattr(sh, 'rotation', 0))
                except Exception:  # noqa: BLE001
                    pos = None
                try:
                    txt = sh.text_frame.text if getattr(sh, 'has_text_frame', False) and sh.has_text_frame else None
                except Exception:  # noqa: BLE001
                    txt = None
                try:
                    kind = str(sh.shape_type)
                except Exception:  # noqa: BLE001  python-pptx cannot classify some placeholders; the same happens on the original
                    kind = type(sh).__name__
                row.append((sh.shape_id, sh.name, kind, pos, txt))
                if kind == '6' or 'GROUP' in kind:  # group
                    walk(sh.shapes)
        walk(s.shapes)
        out.append(row)
    return out


def raw_summary(path):
    """Fallback when python-pptx itself cannot open a deck (it chokes on the deliberately orphaned placeholders in the unsupported-content fixture):
    the same facts read straight from the slide XML: ids, names, frames, rotation, text."""
    ns = {'p': 'http://schemas.openxmlformats.org/presentationml/2006/main', 'a': 'http://schemas.openxmlformats.org/drawingml/2006/main'}
    z = zipfile.ZipFile(path)
    out = []
    for n in sorted(x for x in z.namelist() if x.startswith('ppt/slides/slide') and x.endswith('.xml')):
        root = etree.fromstring(z.read(n))
        row = []
        for c in root.iterfind('.//p:cNvPr', ns):
            sp = c.getparent().getparent()
            xf = sp.find('.//a:xfrm', ns)
            off = xf.find('a:off', ns) if xf is not None else None
            ext = xf.find('a:ext', ns) if xf is not None else None
            row.append((c.get('id'), c.get('name'), etree.QName(sp).localname,
                        (off.get('x'), off.get('y'), ext.get('cx'), ext.get('cy'), xf.get('rot')) if off is not None and ext is not None else None,
                        ''.join(t.text or '' for t in sp.iterfind('.//a:t', ns)) or None))
        out.append(row)
    return out


def lo_pngs(deck, work):
    prof = os.path.join(work, 'p')
    subprocess.run(['soffice', f'-env:UserInstallation=file://{prof}', '--headless', '--convert-to', 'pdf', '--outdir', work, deck], check=True,
                   stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL, timeout=180)
    pdf = os.path.join(work, os.path.splitext(os.path.basename(deck))[0] + '.pdf')
    subprocess.run(['pdftoppm', '-png', '-scale-to-x', str(W), '-scale-to-y', '-1', pdf, os.path.join(work, 'pg')], check=True, timeout=120)
    return sorted(os.path.join(work, f) for f in os.listdir(work) if f.startswith('pg-') and f.endswith('.png'))


def main():
    out_json, pairs = sys.argv[1], sys.argv[2:]
    have_lo = bool(shutil.which('soffice') and shutil.which('pdftoppm'))
    rows = []
    for orig, rt in zip(pairs[0::2], pairs[1::2]):
        name = os.path.splitext(os.path.basename(orig))[0]
        row = {'deck': name, 'errors': []}
        zo, zr = zipfile.ZipFile(orig), zipfile.ZipFile(rt)
        no = [n for n in zo.namelist() if not n.endswith('/')]  # zip folder entries are not package parts
        nr = [n for n in zr.namelist() if not n.endswith('/')]
        row['same_parts'] = sorted(no) == sorted(nr)
        if not row['same_parts']:
            row['errors'].append(f'part lists differ: only in original {sorted(set(no) - set(nr))[:3]}, only in round trip {sorted(set(nr) - set(no))[:3]}')
        changed_bytes, xml_equal, xml_total, untouched_identical, untouched_total = [], 0, 0, 0, 0
        for n in no:
            if n not in nr:
                continue
            a, b = zo.read(n), zr.read(n)
            if a == b:
                untouched_total += 1; untouched_identical += 1
                continue
            if n.endswith('.xml') or n.endswith('.rels'):
                xml_total += 1
                try:
                    if c14n(a) == c14n(b):
                        xml_equal += 1
                    else:
                        row['errors'].append(f'{n}: XML differs after round trip')
                except Exception as e:  # noqa: BLE001
                    row['errors'].append(f'{n}: {str(e)[:80]}')
            else:
                changed_bytes.append(n)
        row.update(parts=len(no), identical_bytes=untouched_identical, reserialised_xml=xml_total, reserialised_equal=xml_equal, binary_changed=changed_bytes)
        if changed_bytes:
            row['errors'].append(f'binary parts changed: {changed_bytes[:3]}')
        try:
            so = pptx_summary(orig)
            try:
                row['pptx_same'] = so == pptx_summary(rt)
            except Exception as e3:  # noqa: BLE001  the original opens but the round trip does not: a real failure
                row['pptx_same'] = False; row['errors'].append('python-pptx opens the original but not the round trip: ' + str(e3)[:80])
        except Exception as e:  # noqa: BLE001
            row['pptx_error'] = 'python-pptx cannot open the original either: ' + str(e)[:100]
            try:
                row['pptx_same'] = raw_summary(orig) == raw_summary(rt); row['pptx_via'] = 'raw XML'
            except Exception as e2:  # noqa: BLE001
                row['pptx_same'] = None; row['pptx_error'] += ' / ' + str(e2)[:80]
        if row['pptx_same'] is False:
            row['errors'].append('python-pptx sees different shapes, positions or text')
        if have_lo:
            with tempfile.TemporaryDirectory() as w1, tempfile.TemporaryDirectory() as w2:
                try:
                    a, b = lo_pngs(os.path.abspath(orig), w1), lo_pngs(os.path.abspath(rt), w2)
                    errs = []
                    for x, y in zip(a, b):
                        ia, ib = Image.open(x).convert('RGB'), Image.open(y).convert('RGB')
                        errs.append(round(float(np.abs(np.asarray(ia, dtype=np.float32) - np.asarray(ib, dtype=np.float32)).mean()), 3))
                    row['lo_slides'] = [len(a), len(b)]; row['lo_error'] = errs
                    if len(a) != len(b) or any(e > 0.5 for e in errs):
                        row['errors'].append(f'LibreOffice renders differ: {errs}')
                except Exception as e:  # noqa: BLE001
                    row['lo_error'] = None; row['errors'].append(f'LibreOffice: {str(e)[:100]}')
        rows.append(row)
        print(json.dumps({k: v for k, v in row.items() if k != 'binary_changed'}))
    json.dump(rows, open(out_json, 'w'))


main()
