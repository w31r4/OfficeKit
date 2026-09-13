## Why

The source-bound theme profile now exposes writable owners for the first five direct RGB accent slots, while accent6 remains observed but source-owned. Closing the final slot gives PPJ a complete, auditable owner for the six direct accent colors without widening the imported theme graph.

## What Changes

- Keep the strict six-color projection from one canonical shared ThemePart.
- Issue a hash-bound `setThemeAccent6Color` capability for `design.theme.accentColors.accent6`.
- On source-preserving export, patch only `a:accent6Color/a:srgbClr/@val`.
- Preserve accent1 through accent5, other color roles, transforms, font scheme, descendants, relationships, and all other package members.
- Reject missing or transformed colors, deletion, capability tampering, unowned edits, and combined theme-field edits.

## Capabilities

### New Capabilities

- `presentation-theme-accent6-color`: Read and source-bound edit of one canonical imported presentation theme accent6 color.

### Modified Capabilities

None.

## Impact

The PPJ theme projection, capability vocabulary, source-bound compiler, and PPTX ThemePart writer gain the final field-qualified accent owner. The existing repeated `PresentationThemeArtifact.accent_rgb` wire field is reused, so no protobuf field or protocol version changes. A focused native regression covers no-op bytes, capability authority, changed-part scope, non-target package preservation, Open XML validity, and second projection.
