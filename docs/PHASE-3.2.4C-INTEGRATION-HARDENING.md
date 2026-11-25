# Phase 3.2.4-C: harden move / resize integration (audit, two confirmed bugs fixed)

An audit of the combined move + resize behaviour (3.2.3-A/B/C and 3.2.4-A/B). Only confirmed bugs were fixed; no new features, no refactoring, no XML change, no export.

## Fixed
1. **The move clamp snapped an off-slide element back in.** The move keeps the element's centre inside the slide, but the allowed range ignored where the drag started. When the centre was already outside (a resize had stretched it there, or the file places it there) the first pixel of any drag threw the element to the nearest edge: a 10 pt drag to the right moved a stretched element **140 pt to the left** and committed that. `clampOffset` now always includes the offset the drag started from: such an element cannot be pushed further out, is not snapped back, and can always be dragged inward. Elements that start inside behave exactly as before. (`js/render/selection.js`)
2. **The preview and the edit-state disagreed about "back at the file frame".** The drag preview treated a frame within 1e-9 pt of the file's as the file's, the edit-state 1e-6 pt. A release between the two left a stretched element on the slide that the edit-state did not hold (a redraw then showed the file frame). Both now use one constant, `SAME_PT` (`js/model/resize-limits.js`).

## Audited and found correct (covered by tests now)
Cancelling a move or a resize (Escape, `pointercancel`, second finger, right button, blur, select, clear) on an element that holds **both** a committed move and a committed resize leaves both records, the version and what is drawn untouched; resize after move and move after resize (composite frames rebuild exactly over `p31-geometry`); moving a resized element back to exactly zero removes only the move record; `reset()` on a selected, moved and resized element leaves the DOM byte-identical to the file drawing; a redraw never changes the picture; `fitText` measures layout sizes, so a restored stretch does not change text fitting. Four seeded random sequences of 45 operations (moves, resizes, cancelled drags, redraws, selections, resets) found no divergence between what is drawn and what the edit-state holds.

## Found, not changed (design, not a defect)
- **Tiny frames cannot be moved by pressing their body.** A press within reach of a visible handle (mouse 7 px, pen 10, touch 16) resizes, and corner handles are never hidden, so on a frame smaller than about twice the reach (e.g. an element resized down to 2 pt) every body press is a resize press. This is the documented 3.2.4-A arbitration; changing it is a behaviour decision.
- **An element larger than four slide sizes snaps down to four on its first resize drag** (the 3.2.4-A maximum applies to the dragged axis from the first pixel). The edit-state accepts what the drag produces.

## Tests
`node tests/selection.mjs`, section "integration hardening (Phase 3.2.4-C)". See TESTING.md.
