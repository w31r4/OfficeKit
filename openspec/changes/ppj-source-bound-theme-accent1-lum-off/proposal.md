## Why

F-15 already lowers authored luminance transforms, and source-bound PPJ can edit exact imported accent1 tint and shade leaves. An imported theme with the same safe direct topology for `lumOff` still keeps that value opaque, leaving one independently auditable luminance field unavailable.

## What Changes

- Project an existing direct `a:accent1Color/a:srgbClr/a:lumOff/@val` as `design.theme.accentTransforms.accent1.lumOff` when its topology is exact.
- Issue a hash-bound `setThemeAccent1LumOff` capability for only that PPJ field.
- On source-preserving export, patch only the existing accent1 luminance offset value in the canonical ThemePart.
- Preserve the base RGB, every other accent and color role, tint/shade, other transforms, font scheme, descendants, relationships, and package members.
- Reject missing, combined, unsupported, deleted, or tampered transform ownership; keep other transform leaves source-owned.

## Capabilities

### New Capabilities

- `presentation-theme-accent1-lum-off`: Read and source-bound edit of one canonical imported presentation theme accent1 luminance offset leaf.

### Modified Capabilities

None.

## Impact

The PPJ theme projection, capability vocabulary, source-bound compiler, and PPTX ThemePart importer/writer gain one field-qualified transform owner. Existing `PresentationThemeArtifact.accent_transforms` and the authored transform wire shape are reused; no protobuf field or protocol version changes. A focused native regression proves no-op bytes, one ThemePart changed-part scope, non-target package preservation, Open XML validity, second projection, and authority rejection.
