## Why

Source-bound PPJ currently keeps an imported theme font scheme opaque, even though the authored model already has a `fontScheme.major` field. F-15 still has a concrete gap for theme font ownership; exposing one existing Latin major typeface gives PPJ a complete, auditable field without rewriting the rest of the theme graph.

## What Changes

- Project the existing `a:majorFont/a:latin/@typeface` as `design.theme.fontScheme.major` when the presentation has one shared ThemePart with stable major and minor Latin typefaces.
- Issue a hash-bound `setThemeFontScheme` capability for `fontScheme.major` only.
- On source-preserving export, patch only the major Latin typeface and report the owning ThemePart as changed.
- Preserve the minor and East Asian/complex-script slots, theme children, relationships, and all other package members; reject missing or ambiguous owners, deletion, capability tampering, and edits to unowned slots.

## Capabilities

### New Capabilities

- `presentation-theme-major-font`: Read and source-bound edit of one canonical imported presentation theme major Latin typeface.

### Modified Capabilities

None.

## Impact

The PPJ theme projection, schema capability vocabulary, source-bound compiler, and PPTX ThemePart writer gain this narrow field. The existing `PresentationThemeArtifact.major_font_family` wire field is reused, so no new wire number or protocol version is needed. A focused native regression will cover no-op bytes, capability authority, changed-part scope, non-target package preservation, Open XML validity, and second projection.
