## Why

The source-bound theme profile currently exposes the six imported accent values but grants a write capability only for accentColors.accent1. F-15 still leaves the next direct theme color slot source-owned; closing accentColors.accent2 gives PPJ one more complete, auditable field without widening the theme graph.

## What Changes

- Keep the strict six-color projection from one canonical shared ThemePart.
- Issue a hash-bound setThemeAccent2Color capability for design.theme.accentColors.accent2.
- On source-preserving export, patch only a:accent2Color/a:srgbClr/@val.
- Preserve accent1 and accent3..6, other color roles, transforms, font scheme, descendants, relationships, and all other package members.
- Reject missing or transformed colors, deletion, capability tampering, unowned accent edits, and combined theme-field edits.

## Capabilities

### New Capabilities

- presentation-theme-accent2-color: Read and source-bound edit of one canonical imported presentation theme accent2 color.

### Modified Capabilities

None.

## Impact

The PPJ theme projection, capability vocabulary, source-bound compiler, and PPTX ThemePart writer gain one field-qualified owner. The existing repeated PresentationThemeArtifact.accent_rgb wire field is reused, so no protobuf field or protocol version changes. A focused native regression covers no-op bytes, capability authority, changed-part scope, non-target package preservation, Open XML validity, and second projection.
