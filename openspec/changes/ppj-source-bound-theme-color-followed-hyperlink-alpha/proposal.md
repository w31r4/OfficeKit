## Why

F-15 now has source-bound owners for the imported dark1, light1, dark2, light2, and hyperlink roles, but the direct RGB followedHyperlink role remains source-owned even when it is structurally simple. The existing canonical ThemePart boundary can expose followedHyperlink as one independently authorized PPJ field without rebuilding the theme graph.

## What Changes

- Project an existing direct `a:folHlink/a:srgbClr/@val` with an existing `a:alpha/@val` when present as `design.theme.colorRoles.followedHyperlink` alongside the existing six-role observation.
- Issue a hash-bound `setThemeColorRoleFollowedHyperlink` capability for that field.
- On source-preserving export, patch only the existing followedHyperlink RGB value and its existing alpha value in the canonical ThemePart.
- Preserve dark1, light1, dark2, light2, hyperlink, accents, transforms, font scheme, descendants, relationships, and package members.
- Reject missing, scheme-colored, transformed, alpha-presence-changing, ambiguous, deleted, or tampered ownership.

## Capabilities

### Modified Capabilities

- `presentation-theme-color-role-followedHyperlink`: the existing source-bound field now covers a strict direct RGBA followedHyperlink owner in addition to direct RGB, preserving the existing alpha child.

## Impact

The PPJ theme projection, capability vocabulary, source-bound compiler, and PPTX ThemePart importer/writer gain one field-qualified color-role owner. Existing `PresentationThemeArtifact.followedHyperlink_rgb` is reused for the canonical RGBA string; no protobuf field or protocol version changes. A focused native regression proves no-op bytes, one ThemePart changed-part scope, non-target package preservation, Open XML validity, second projection, and authority rejection.
