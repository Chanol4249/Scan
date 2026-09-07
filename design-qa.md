# Design QA — TRACK home card redesign

## Evidence

- Source visual truth: `/workspace/scratch/3231f7436500/upload/IMG_0378.png`
- Browser-rendered implementation: `/workspace/scratch/scan-form-qa-mobile-final.png`
- Normalized implementation crop: `/workspace/scratch/3231f7436500/qa-evidence/qa-mobile-implementation.png`
- Side-by-side comparison: `/workspace/scratch/3231f7436500/qa-evidence/qa-comparison.png`
- State: authenticated home screen with three form choices
- CSS viewport: `393 × 852 px` in the mobile preview frame
- Browser screenshot: `1363 × 936 px`; implementation was cropped to the `393 × 852 px` frame
- Source pixels: `1320 × 2868 px`; normalized to `393 × 852 px`
- Density normalization: both comparison panels are `393 × 852 px`

## Full-view comparison

The source clearly shows the problem state: TRACK occupies only the left grid column, leaves an empty right column, uses a tiny clipboard emoji, and creates excessive vertical scrolling. The revised browser render keeps ADJ and RTC as the two equal cards in the first row, then spans TRACK across both columns as a shorter horizontal card. The new generated illustration is large, sharp, and uses the same soft mint/green 3D art direction as the other two cards.

## Required fidelity surfaces

- Fonts and typography: existing Thai/Latin font stack, weights, hierarchy, wrapping, and button labels are preserved. TRACK copy remains legible at 393 px without truncation.
- Spacing and layout rhythm: the first-row grid is unchanged; TRACK now spans `1 / -1`, uses a balanced 43/57 image-copy split, and removes the empty-column imbalance.
- Colors and visual tokens: existing navy, mint, green CTA gradient, border, radius, and shadow tokens are reused.
- Image quality and asset fidelity: the emoji placeholder was replaced with a production WebP asset at `480 × 480 px`; crop, contrast, and focal density are appropriate for the card slot.
- Copy and content: ADJ, RTC, TRACK, Thai names, English names, and CTA copy are unchanged.

Focused-region comparison was not needed: the only redesigned component is the TRACK card, and its illustration, typography, spacing, and CTA are clearly readable in the normalized full-view comparison.

## Interaction and runtime checks

- TRACK card opens `Tracking Receipt Recording`.
- `+ เพิ่มรายการ` opens the TRACK form.
- PO entry, completed-status selection, and save produce a visible `PO PO-9` list item.
- No application console errors were observed. Logged errors came only from the cloud-browser extension and not from the app origin.

## Findings and comparison history

### Iteration 1 — source problem

- P1: TRACK was trapped in the left grid column and left half the row empty.
- P1: the clipboard emoji was too small and did not match the production 3D image system.
- P2: the third card repeated the tall portrait proportion and created unnecessary scrolling.

### Fixes applied

- Spanned TRACK across both grid columns.
- Changed TRACK to a compact horizontal layout with a responsive mobile rule.
- Replaced the emoji with a generated 3D purchase-order, delivery, calendar, and status-check illustration.

### Post-fix evidence

- The normalized browser render shows an even two-card first row and a full-width horizontal TRACK card beneath it.
- At 393 px, all text and the CTA fit without clipping or horizontal overflow.
- Primary TRACK interactions complete successfully.

No actionable P0, P1, or P2 findings remain. No P3 polish items are required for this pass.

final result: passed
