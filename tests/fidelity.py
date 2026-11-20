#!/usr/bin/env python3
"""Visual fidelity check for the Rendered tab (Phase 3.1).

Renders each deck with LibreOffice (an independent PPTX renderer) and compares it, slide by slide, with the PNG that the browser produced from the
Rendered tab (tests/render.mjs saves them as <deck>-<n>.png, 960 px wide). Both images are reduced to a coarse grid, so anti-aliasing and font hinting do not
matter, but a missing shape, a wrong colour, a shifted or mis-sized element, flipped / rotated content or lost text blocks do.

usage: fidelity.py OUT_JSON SHOTS_DIR DECK.pptx [DECK.pptx ...]
Prints one JSON line per slide and writes the list to OUT_JSON. Exit code 0 always (the caller decides with thresholds); exit 3 when soffice is missing.
"""
import json, os, shutil, subprocess, sys, tempfile
from PIL import Image, ImageFilter
import numpy as np

W = 960
GRID = (96, 54)


def lo_render(deck, work):
    prof = os.path.join(work, 'profile')
    subprocess.run(['soffice', f'-env:UserInstallation=file://{prof}', '--headless', '--convert-to', 'pdf', '--outdir', work, deck],
                   check=True, stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL, timeout=180)
    pdf = os.path.join(work, os.path.splitext(os.path.basename(deck))[0] + '.pdf')
    base = os.path.join(work, 'lo')
    subprocess.run(['pdftoppm', '-png', '-scale-to-x', str(W), '-scale-to-y', '-1', pdf, base], check=True, timeout=120)
    files = sorted(f for f in os.listdir(work) if f.startswith('lo-') and f.endswith('.png'))
    return [os.path.join(work, f) for f in files]


def feat(path, size):
    im = Image.open(path).convert('RGB')
    if im.size != size:
        im = im.resize(size, Image.LANCZOS)
    return im


def compare(a, b):
    ga = np.asarray(a.resize(GRID, Image.BOX), dtype=np.float32)
    gb = np.asarray(b.resize(GRID, Image.BOX), dtype=np.float32)
    d = np.abs(ga - gb).mean(axis=2)  # per-tile mean absolute colour difference
    # "ink" maps: where the picture is not the page colour (finds missing / extra blocks regardless of colour shifts)
    la = np.asarray(a.convert('L').resize(GRID, Image.BOX), dtype=np.float32)
    lb = np.asarray(b.convert('L').resize(GRID, Image.BOX), dtype=np.float32)
    bga, bgb = np.median(la), np.median(lb)
    ia, ib = np.abs(la - bga) > 24, np.abs(lb - bgb) > 24
    inter = np.logical_and(ia, ib).sum()
    union = np.logical_or(ia, ib).sum()
    return {
        'mean': round(float(d.mean()), 2),                  # average colour error over the slide, 0..255
        'tiles_over_40': round(float((d > 40).mean()), 4),   # share of the slide that differs strongly
        'ink_iou': round(float(inter / union), 3) if union else 1.0,  # overlap of "non-background" areas
    }


def main():
    out_json, shots, decks = sys.argv[1], sys.argv[2], sys.argv[3:]
    if not shutil.which('soffice') or not shutil.which('pdftoppm'):
        print('SKIP: soffice / pdftoppm not available', file=sys.stderr)
        sys.exit(3)
    rows = []
    for deck in decks:
        name = os.path.splitext(os.path.basename(deck))[0]
        with tempfile.TemporaryDirectory() as work:
            try:
                refs = lo_render(os.path.abspath(deck), work)
            except Exception as e:  # noqa: BLE001
                rows.append({'deck': name, 'slide': 0, 'error': str(e)[:200]})
                continue
            for i, ref in enumerate(refs, 1):
                ours = os.path.join(shots, f'{name}-{i}.png')
                if not os.path.exists(ours):
                    rows.append({'deck': name, 'slide': i, 'error': 'no browser screenshot'})
                    continue
                a = Image.open(ours).convert('RGB')
                b = feat(ref, a.size)
                row = {'deck': name, 'slide': i, 'size': list(a.size), 'lo_size': list(Image.open(ref).size), **compare(a, b)}
                rows.append(row)
                print(json.dumps(row))
    json.dump(rows, open(out_json, 'w'))


main()
