# Phase 3.1: rendering engine and the read-only "Rendered" tab

Draws existing PPTX slides as native, editable-style elements (real text, SVG vector shapes, `<table>`, clipped pictures). Never rasterized. Read-only: nothing is saved or written back. Phase 3.2 (editing) is not started.

## Architecture

| Layer | File | Role |
|---|---|---|
| Scene model | `js/model/scene.js` | `buildScene(parsed, i)`: derived, in-memory, read-only view of one slide with live references into the original OOXML DOM (the DOM stays the single source of truth). Stable IDs `part#cNvPr-id` (`~N` duplicates, `@path` missing id, `~choiceN`/`~fallback` branches). `z` = document order. Affine matrices through groups. Placeholder inheritance slide → layout (idx, then type) → master (type). Every element that cannot be placed is kept in `scene.unplaced`. |
| Style resolution | `js/model/resolve.js` | Fill, line, shadow, glow, text bodies/paragraphs/runs, table styles, backgrounds. Reuses the exported Phase 2 helpers in `js/pptx/design.js` (only `export` keywords added; logic and analysis JSON unchanged). |
| Geometry | `js/render/geometry.js` | ~55 preset shapes with adjust values, spec-literal `custGeom` (guides, arcTo, all path commands). |
| Renderer | `js/render/render.js` | `renderSlide(scene, host, {width})` → `{root, nodes, report(), setWidth(px), fitText(), dispose()}`. Strict CSP: no inline styles and no eval; all styling through the CSSOM / SVG attributes and one CSS variable `--u` (px per pt), so resizing is one variable change. |
| Fonts | `js/render/fonts.js`, `css/fonts.css`, `fonts/` | The deck's own font is tried first, then a metric-compatible self-hosted substitute, then Arabic-capable fallbacks, then a generic family. |
| UI | `js/ui/rendered.js`, `css/render.css` | Third tab "Rendered": filmstrip, slide, honest support summary (counts, per-element reasons, fonts used). Selecting a slide in the Slides tab changes the rendered slide. Lazy: nothing is built until the tab is opened. |

Existing tabs, routes, Firebase, auth, Firestore data and `package.json` are untouched. `js/ui/preview.js` has a surgical hook (17 changed lines).

## Fonts

Self-hosted in `fonts/`, all unmodified, each with its licence in `fonts/licenses/` and sha256 sums in `fonts/README.md`:
Carlito (OFL, Calibri metrics), Liberation Sans/Serif/Mono (OFL, Arial/Times/Courier metrics), Caladea (Apache-2.0, Cambria metrics), DejaVu Sans (Bitstream Vera licence, Arabic coverage).
Families are exposed as `SF Carlito` etc. so they never shadow a font that is installed. Original font choices are preserved: the font named in the file is always first in the stack.

**Adding an Arabic font** (recommended: Noto Naskh Arabic or Noto Sans Arabic, both OFL): put the TTF/WOFF2 in `fonts/`, add a licence in `fonts/licenses/`, add an `@font-face` in `css/fonts.css` named `Noto Naskh Arabic`, and list the file in `fonts/README.md`. The fallback stack already looks for those names. Nothing else changes.

## Tests

- `npm test` (108) and `npm run test:analysis` (206): unchanged Phase 2.5 suites. Design JSON hashes identical to baseline (only the new p31 fixtures were added).
- `node tests/render.mjs` (670 checks): headless Chromium against the app served with the real `vercel.json` headers (so CSP violations fail the run). Sections: geometry, nothing-dropped, XML immutability, unsupported content, layers, style resolution, renderer DOM, fonts (CDP `getPlatformFontsForNode` verifies which face is actually drawn), Rendered tab UI flow, placeholder-inheritance regression, fidelity vs LibreOffice, round trip.
- Fidelity: `tests/fidelity.py` compares each browser render with an independent LibreOffice render (coarse grid, thresholds in the test). Average error 1.64/255 over 33 slides.
- Round trip: scene + render every slide, re-serialise the live XML into a new .pptx; `tests/roundtrip.py` checks identical package parts, byte-identical untouched parts (media, charts, SmartArt, OLE…), C14N-equal XML, same shapes via python-pptx, identical LibreOffice render.
- Python tools (lxml, python-pptx, Pillow, numpy, LibreOffice + pdftoppm) are optional; those checks print SKIPPED when missing.
- Real decks: `tests/real/*.pptx` (see its README). `REAL_PPTX=/path/a.pptx:/path/b.pptx node tests/render.mjs` adds your own decks. `SHOTS=/some/dir` keeps screenshots, round-trip files and JSON reports.
- Regenerate generated fixtures: `python3 -I tests/make_fixture_p31.py`.

## Support matrix

**Native** (drawn from the file with no approximation): rectangle and ~55 preset shapes with adjust handles; custom geometry; solid and linear-gradient fills; outlines with width, dash, cap, join, arrowheads; outer shadow; pictures with crop, rotation, flip and clip geometry; groups with nested transforms; text runs and paragraphs (bold/italic/underline/strike, size, colour, spacing, alignment, indents, bullets and numbering, RTL, columns, vertical text); tables with merged cells, row/column sizes, borders and stored table styles; theme colours and fonts; placeholder inheritance (slide → layout → master); slide backgrounds; master/layout shapes and hidden-master-shape handling; hidden elements (kept, not shown).

**Approximated** (drawn, but marked with a reason in the tab): pattern fills (flat mix); path gradients (radial approximation); glow (double drop-shadow halo); simplified presets (heart, cloud, sun, lightning bolt, rectangular callout) and unknown presets (drawn as a rectangle); 3D, bevel, reflection, soft edges; colour modifiers beyond the common set; text effects; picture effects; tiled fills; "resize shape to fit text" (stored frame kept); text "shrink on overflow" (stored `fontScale` applied, or measured in the browser); skewed shapes; built-in table styles not stored in the file (Medium Style 2 only); picture bullets / symbol bullets without a Unicode equivalent; the AlternateContent wrapper; audio/video (poster only).

**Opaque** (kept in the file and in the model, shown as a labelled box or listed, never dropped): charts; SmartArt; OLE objects; ink and other content parts; unknown vendor elements; unplaced placeholders (no resolvable position); the AlternateContent branch not drawn; EMF/WMF/TIFF and linked pictures (placeholder box).

## Known limitations

- No genuine PowerPoint-, Keynote- or Google-exported decks were available in the build environment. Real decks used: PowerPoint-authored default masters (English, Arabic RTL) and LibreOffice-saved decks, plus generated fixtures. Please run `REAL_PPTX=...` on your own decks (especially Arabic ones).
- Arabic `algn` + `rtl` alignment semantics were checked against LibreOffice, not against real PowerPoint.
- Arabic text falls back to DejaVu Sans (Noto Arabic could not be fetched in the build sandbox). Readable and correctly shaped, but not typographically ideal. See "Adding an Arabic font".
- Phase 2 `phShape` (design.js) matches placeholders by idx on both slide→layout and layout→master hops; the master hop should match by type. Left unchanged to keep analysis output identical. The renderer uses its own correct `phSources`.
