## Why

F-15 now has source-bound owners for the imported dark1 and light1 roles, but the direct RGB dark2 role remains source-owned even when it is structurally simple. The existing canonical ThemePart boundary can expose dark2 as one independently authorized PPJ field without rebuilding the theme graph.

## What Changes

- Project an existing direct `a:dk2/a:srgbClr/@val` as `design.theme.colorRoles.dark2` alongside the existing six-role observation.
- Issue a hash-bound `setThemeColorRoleDark2` capability for that field.
- On source-preserving export, patch only the existing dark2 RGB value in the canonical ThemePart.
- Preserve dark1, light1, light2, hyperlink, followedHyperlink, accents, transforms, font scheme, descendants, relationships, and package members.
- Reject missing, scheme-colored, transformed, alpha-bearing, ambiguous, deleted, or tampered ownership.

## Capabilities

### New Capabilities

- `presentation-theme-color-role-dark2`: Read and source-bound edit of one canonical imported presentation theme dark2 color role.

### Modified Capabilities

None.

## Impact

The PPJ theme projection, capability vocabulary, source-bound compiler, and PPTX ThemePart importer/writer gain one field-qualified color-role owner. Existing `PresentationThemeArtifact.dark2_rgb` is reused; no protobuf field or protocol version changes. A focused native regression proves no-op bytes, one ThemePart changed-part scope, non-target package preservation, Open XML validity, second projection, and authority rejection.
