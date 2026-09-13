## Why

The source-bound theme profile now exposes writable owners for the first two direct RGB accent slots, while accent3 remains observed but source-owned. Closing the next slot gives PPJ one more complete, auditable theme field without widening the imported theme graph.

## What Changes

- Keep the strict six-color projection from one canonical shared ThemePart.
- Issue a hash-bound `setThemeAccent3Color` capability for `design.theme.accentColors.accent3`.
- On source-preserving export, patch only `a:accent3Color/a:srgbClr/@val`.
- Preserve accent1, accent2, accent4..6, other color roles, transforms, font scheme, descendants, relationships, and all other package members.
- Reject missing or transformed colors, deletion, capability tampering, unowned accent edits, and combined theme-field edits.

## Capabilities

### New Capabilities

- `presentation-theme-accent3-color`: Read and source-bound edit of one canonical imported presentation theme accent3 color.

### Modified Capabilities

None.

## Impact

The PPJ theme projection, capability vocabulary, source-bound compiler, and PPTX ThemePart writer gain one field-qualified owner. The existing repeated `PresentationThemeArtifact.accent_rgb` wire field is reused, so no protobuf field or protocol version changes. A focused native regression covers no-op bytes, capability authority, changed-part scope, non-target package preservation, Open XML validity, and second projection.
