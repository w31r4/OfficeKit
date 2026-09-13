## Why

F-15 now has source-bound owners for the imported dark1, light1, dark2, and light2 roles, but a simple direct RGB hyperlink role remains source-owned. The canonical ThemePart boundary can expose this role as one independently authorized PPJ field without rebuilding the theme graph.

## What Changes

- Project an existing direct `a:hlink/a:srgbClr/@val` as `design.theme.colorRoles.hyperlink` alongside the existing six-role observation.
- Issue a hash-bound `setThemeColorRoleHyperlink` capability for that field.
- On source-preserving export, patch only the existing hyperlink RGB value in the canonical ThemePart.
- Preserve dark1, light1, dark2, light2, followedHyperlink, accents, transforms, font scheme, descendants, relationships, and package members.
- Reject missing, scheme-colored, transformed, alpha-bearing, ambiguous, deleted, or tampered ownership.

## Capabilities

### New Capabilities

- `presentation-theme-color-role-hyperlink`: Read and source-bound edit of one canonical imported presentation theme hyperlink color role.

### Modified Capabilities

None.

## Impact

The PPJ theme projection, capability vocabulary, source-bound compiler, and PPTX ThemePart importer/writer gain one field-qualified color-role owner. Existing `PresentationThemeArtifact.hyperlink_rgb` is reused; no protobuf field or protocol version changes. A focused native regression proves no-op bytes, one ThemePart changed-part scope, non-target package preservation, Open XML validity, second projection, and authority rejection.
