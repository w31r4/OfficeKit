## Why

F-15 already has an authored `design.theme.colorRoles` model and a source-bound dark1 owner, but the imported light1 role remains source-owned even when it is a simple direct RGB leaf. The same canonical ThemePart boundary can expose light1 as one independently authorized PPJ field without rebuilding the theme graph.

## What Changes

- Project an existing direct `a:lt1/a:srgbClr/@val` with an existing `a:alpha/@val` when present as `design.theme.colorRoles.light1` alongside the existing six-role observation.
- Issue a hash-bound `setThemeColorRoleLight1` capability for that field.
- On source-preserving export, patch only the existing light1 RGB value and its existing alpha value in the canonical ThemePart.
- Preserve dark1, dark2, light2, hyperlink, followedHyperlink, accents, transforms, font scheme, descendants, relationships, and package members.
- Reject missing, scheme-colored, transformed, alpha-presence-changing, ambiguous, deleted, or tampered ownership.

## Capabilities

### Modified Capabilities

- `presentation-theme-color-role-light1`: the existing source-bound field now covers a strict direct RGBA light1 owner in addition to direct RGB, preserving the existing alpha child.

## Impact

The PPJ theme projection, capability vocabulary, source-bound compiler, and PPTX ThemePart importer/writer gain one field-qualified color-role owner. Existing `PresentationThemeArtifact.light1_rgb` is reused for the canonical RGBA string; no protobuf field or protocol version changes. A focused native regression proves no-op bytes, one ThemePart changed-part scope, non-target package preservation, Open XML validity, second projection, and authority rejection.
