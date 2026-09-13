## Why

F-15 already lowers authored accent transforms, but an imported theme with a single direct `a:tint` on accent1 remains entirely source-owned. That leaves a safe, field-qualified transform leaf unavailable even though the canonical ThemePart can preserve it without rebuilding the theme graph.

## What Changes

- Project an existing direct `a:accent1Color/a:srgbClr/a:tint/@val` as `design.theme.accentTransforms.accent1.tint` when its topology is exact.
- Issue a hash-bound `setThemeAccent1Tint` capability for only that PPJ field.
- On source-preserving export, patch only the existing accent1 tint value in the canonical ThemePart.
- Preserve the accent1 base RGB, every other accent and color role, other transforms, font scheme, descendants, relationships, and package members.
- Reject missing, combined, unsupported, deleted, or tampered transform ownership; keep other transform leaves source-owned.

## Capabilities

### New Capabilities

- `presentation-theme-accent1-tint`: Read and source-bound edit of one canonical imported presentation theme accent1 tint leaf.

### Modified Capabilities

None.

## Impact

The PPJ theme projection, capability vocabulary, source-bound compiler, and PPTX ThemePart importer/writer gain one field-qualified transform owner. Existing `PresentationThemeArtifact.accent_transforms` and the authored transform wire shape are reused; no protobuf field or protocol version changes. A focused native regression proves no-op bytes, one ThemePart changed-part scope, non-target package preservation, Open XML validity, second projection, and authority rejection.
