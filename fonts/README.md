# Self-hosted fonts (SlideFlow Phase 3.1)

These are the **original, unmodified** font files, shipped only so the Rendered tab can show a deck with the same line breaks on machines that do not have the fonts the deck asks for. The deck's own font (Calibri, Arial...) is always tried first; these are the fallback.

| Family used on the page | Stands in for | Upstream | Licence | Files |
|---|---|---|---|---|
| SF Carlito | Calibri | Carlito 1.x (tyPoland, Lato-based), Debian `fonts-crosextra-carlito` 20230309-2 | SIL OFL 1.1 | Regular, Bold, Italic, BoldItalic |
| SF Caladea | Cambria | Caladea (Huerta Tipografica / Google crosextrafonts), Debian `fonts-crosextra-caladea` 20130214-2.1 | Apache 2.0 | Regular, Bold, Italic, BoldItalic |
| SF Liberation Sans | Arial, Helvetica, Tahoma, Verdana, Segoe UI | Liberation Fonts 2.1.5, Debian `fonts-liberation2` | SIL OFL 1.1 | Regular, Bold, Italic, BoldItalic |
| SF Liberation Serif | Times New Roman, Georgia | Liberation Fonts 2.1.5 | SIL OFL 1.1 | Regular, Bold, Italic, BoldItalic |
| SF Liberation Mono | Courier New, Consolas | Liberation Fonts 2.1.5 | SIL OFL 1.1 | Regular, Bold |
| SF DejaVu Sans | Arabic text (and any glyph the other faces lack) | DejaVu Fonts 2.37, Debian `fonts-dejavu-core` 2.37-8 | Bitstream Vera licence + public-domain changes | Regular, Bold |

## Rules

- Files are never edited, subset, converted or renamed in content (the OFL Reserved Font Names stay valid because nothing is modified). The family names used on the page carry an `SF ` prefix only in CSS.
- Licence texts are in `fonts/licenses/`. Keep them next to the files if you move the fonts.
- Only fonts whose licence allows redistribution may be added here. Do not add Calibri, Arial, Times New Roman, Cambria or any other commercial font.
- Arabic: `SF DejaVu Sans` is the guaranteed Arabic-capable face. For nicer Arabic typography add an OFL font such as Noto Naskh Arabic or Amiri: drop the TTFs here with their licence, add `@font-face` blocks to `css/fonts.css`, and list the family in `ARABIC_FALLBACK` in `js/render/fonts.js`. Nothing else changes.
- The files load lazily: a face is downloaded only when a slide uses it and the original font is not installed.

## Checksums (sha256)

```
f6418f708baede9789daef5d458c0f53d2a888af9820e8062934e504fedc6595  Carlito-Regular.ttf  (613 KB)
bb5d20f79b82599ec72983597437373a80f2d2085fa91fc144fd74e876a594db  Carlito-Bold.ttf  (666 KB)
0b019225e58d702bfedcbd35c21696769f8ee115cb6343f84c2f240312450d1c  Carlito-Italic.ttf  (600 KB)
b32928186c119599e03ca6a1ffc680fdcb7fac95772f4b95d989cf6cd3861517  Carlito-BoldItalic.ttf  (789 KB)
d2f6cad33f191e65b68bd74e6d4f7708080a41b32db635866109df3090265d91  Caladea-Regular.ttf  (57 KB)
74eda4fc5ffba0d8dc4aa76d41499f5ab76168d9ed6141c6417064c6244db9b6  Caladea-Bold.ttf  (57 KB)
9c968bf60ba1e851cdfea77c71e3540c57792f77482dd241acc29d1425569c4e  Caladea-Italic.ttf  (60 KB)
f47a35ad6cd0efa9914d93f9d03c93e30c55da43674bccfabac73b3c0d522c01  Caladea-BoldItalic.ttf  (60 KB)
4659bc0c58c5028dd488ec928d41d9265db43d9b669fc14ca8b0832daca7b144  LiberationSans-Regular.ttf  (401 KB)
3973aa5054fb467dd5627245d3dc82e37bf16fe075756156a570455871351582  LiberationSans-Bold.ttf  (404 KB)
830c5fa600505fb4c1a271b4271c53c44bae43f492b2a240d0e98a3a7a380121  LiberationSans-Italic.ttf  (406 KB)
c80fa7f2ffa0e01d4d8dcd6a6d1e43eda665222d1b4db597dde1174c456006cf  LiberationSans-BoldItalic.ttf  (399 KB)
705903ae1382f0150f115d6f10a30f5b2f3d6ade649c474afe6366c39174fb65  LiberationSerif-Regular.ttf  (384 KB)
9e66c25e20868756bf12eb24a520a4552727c9cd108f577f7ae33e3c0110e39c  LiberationSerif-Bold.ttf  (361 KB)
eaeaf7f2b12544ecde64290c1fa7ec403ff70fd58b313637b265400b20249242  LiberationSerif-Italic.ttf  (366 KB)
379010e87421a883f7bbfa7936d23dfe0257a54ce850819b74556e9f2c615a8a  LiberationSerif-BoldItalic.ttf  (368 KB)
395fa5ab8d40c8eba390ced528744ea75a7f69aabf3e68b6f925ca0e39a27370  LiberationMono-Regular.ttf  (312 KB)
626655e94dd82f3f42549daf995c921b0915fa8ab1f4b839559e8892ea41d240  LiberationMono-Bold.ttf  (300 KB)
ae7b7855e115a5966d8b1b3f80f254ccc117ec86f9965e202ee2940453837280  DejaVuSans.ttf  (741 KB)
5c1247acef7f2b8522a31742c76d6adcb5569bacc0be7ceaa4dc39dd252ce895  DejaVuSans-Bold.ttf  (692 KB)
```
