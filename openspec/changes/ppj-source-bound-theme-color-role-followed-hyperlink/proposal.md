## Why

F-15 now has a source-bound owner for the imported hyperlink role, but the matching `followedHyperlink` role remains source-owned even when it is a direct RGB leaf. The canonical ThemePart boundary can expose this last color-role leaf as one independently authorized PPJ field without rebuilding the theme graph.

## What Changes

- Project an existing direct `a:folHlink/a:srgbClr/@val` as `design.theme.colorRoles.followedHyperlink` alongside the existing six-role observation.
- Issue a hash-bound `setThemeColorRoleFollowedHyperlink` capability for that field.
- On source-preserving export, patch only the existing followed-hyperlink RGB value in the canonical ThemePart.
- Preserve dark1, light1, dark2, light2, hyperlink, accents, transforms, font scheme, descendants, relationships, and package members.
- Reject missing, scheme-colored, transformed, alpha-bearing, ambiguous, deleted, or tampered ownership.

## Capabilities

### New Capabilities

- `presentation-theme-color-role-followed-hyperlink`: Read and source-bound edit of one canonical imported presentation theme followed-hyperlink color role.

### Modified Capabilities

None.

## Impact

The PPJ theme projection, capability vocabulary, source-bound compiler, and PPTX ThemePart importer/writer gain one field-qualified color-role owner. Existing `PresentationThemeArtifact.followed_hyperlink_rgb` is reused; no protobuf field or protocol version changes. A focused native regression proves no-op bytes, one ThemePart changed-part scope, non-target package preservation, Open XML validity, second projection, and authority rejection.
