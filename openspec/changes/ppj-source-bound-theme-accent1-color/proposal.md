## Why

Source-bound PPJ currently leaves imported theme colors opaque even though the authored model already has six accent fields. F-15 still has a concrete gap for one theme color owner; exposing one existing direct RGB accent gives PPJ a complete, auditable field without rewriting the remaining theme graph.

## What Changes

- Project six existing direct RGB accent nodes as `design.theme.accentColors` when the presentation has one shared ThemePart and all six accents are strict direct colors.
- Issue a hash-bound `setThemeAccent1Color` capability for `accentColors.accent1` only.
- On source-preserving export, patch only `a:accent1Color/a:srgbClr/@val` and report the owning ThemePart as changed.
- Preserve the other five accents, color roles, transforms, font scheme, descendants, relationships, and all other package members; reject transformed or missing colors, deletion, capability tampering, and edits to unowned slots.

## Capabilities

### New Capabilities

- `presentation-theme-accent1-color`: Read and source-bound edit of one canonical imported presentation theme accent color.

### Modified Capabilities

None.

## Impact

The PPJ theme projection, schema capability vocabulary, source-bound compiler, and PPTX ThemePart writer gain this narrow field. The existing repeated `PresentationThemeArtifact.accent_rgb` wire field is reused, so no new wire number or protocol version is needed. A focused native regression covers no-op bytes, capability authority, changed-part scope, non-target package preservation, Open XML validity, and second projection.
