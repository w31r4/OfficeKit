## Why

F-15 already lowers authored alpha transforms, while source-bound PPJ currently keeps a direct imported `a:alphaMod` leaf opaque. Exposing one strict alpha-modulation owner closes a concrete theme field without pretending to edit the surrounding effect or inheritance graph.

## What Changes

- Project an existing direct `a:accent1Color/a:srgbClr/a:alphaMod/@val` as `design.theme.accentTransforms.accent1.alphaMod` when its topology is exact.
- Issue a hash-bound `setThemeAccent1AlphaMod` capability for only that PPJ field.
- On source-preserving export, patch only the existing accent1 alpha-modulation value in the canonical ThemePart.
- Preserve the base RGB, every other accent and color role, tint/shade/luminance transforms, other alpha and color transforms, font scheme, descendants, relationships, and package members.
- Reject missing, combined, unsupported, deleted, out-of-range, or tampered transform ownership; keep other transform leaves source-owned.

## Capabilities

### New Capabilities

- `presentation-theme-accent1-alpha-mod`: Read and source-bound edit of one canonical imported presentation theme accent1 alpha-modulation leaf.

### Modified Capabilities

None.

## Impact

The PPJ theme projection, capability vocabulary, source-bound compiler, and PPTX ThemePart importer/writer gain one field-qualified transform owner. Existing `PresentationThemeColorTransform.AlphaModulationThousandth` and the authored transform wire shape are reused; no protobuf field or protocol version changes. A focused native regression proves no-op bytes, one ThemePart changed-part scope, non-target package preservation, Open XML validity, second projection, and authority rejection.
