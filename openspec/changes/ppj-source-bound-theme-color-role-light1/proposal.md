## Why

F-15 already has an authored `design.theme.colorRoles` model and a source-bound dark1 owner, but the imported light1 role remains source-owned even when it is a simple direct RGB leaf. The same canonical ThemePart boundary can expose light1 as one independently authorized PPJ field without rebuilding the theme graph.

## What Changes

- Project an existing direct `a:lt1/a:srgbClr/@val` as `design.theme.colorRoles.light1` alongside the existing six-role observation.
- Issue a hash-bound `setThemeColorRoleLight1` capability for that field.
- On source-preserving export, patch only the existing light1 RGB value in the canonical ThemePart.
- Preserve dark1, dark2, light2, hyperlink, followedHyperlink, accents, transforms, font scheme, descendants, relationships, and package members.
- Reject missing, scheme-colored, transformed, alpha-bearing, ambiguous, deleted, or tampered ownership.

## Capabilities

### New Capabilities

- `presentation-theme-color-role-light1`: Read and source-bound edit of one canonical imported presentation theme light1 color role.

### Modified Capabilities

None.

## Impact

The PPJ theme projection, capability vocabulary, source-bound compiler, and PPTX ThemePart importer/writer gain one field-qualified color-role owner. Existing `PresentationThemeArtifact.light1_rgb` is reused; no protobuf field or protocol version changes. A focused native regression proves no-op bytes, one ThemePart changed-part scope, non-target package preservation, Open XML validity, second projection, and authority rejection.
