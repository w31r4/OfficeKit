## Why

F-15 already lowers authored `design.theme.colorRoles`, but imported theme color roles remain source-owned. The direct `a:dk1` role is a bounded single-leaf candidate that can be exposed without rebuilding the theme graph.

## What Changes

- Project an existing direct RGB/RGBA `a:dk1/a:srgbClr/@val` with its existing `a:alpha/@val` when present as `design.theme.colorRoles.dark1`.
- Issue a hash-bound `setThemeColorRoleDark1` capability for that field.
- On source-preserving export, patch only the existing dark1 RGB value and its existing alpha value in the canonical ThemePart.
- Preserve all other color roles, accents, transforms, font scheme, descendants, relationships, and package members.
- Reject missing, scheme-colored, transformed, alpha-presence-changing, ambiguous, deleted, or tampered ownership.

## Capabilities

### Modified Capabilities

- `presentation-theme-color-role-dark1`: the existing source-bound field now covers a strict direct RGBA dark1 owner in addition to direct RGB, preserving the existing alpha child.

## Impact

The PPJ theme projection, capability vocabulary, source-bound compiler, and PPTX ThemePart importer/writer gain one field-qualified color-role owner. Existing `PresentationThemeArtifact.dark1_rgb` is reused for the canonical RGBA string; no protobuf field or protocol version changes. A focused native regression proves no-op bytes, one ThemePart changed-part scope, non-target package preservation, Open XML validity, second projection, and authority rejection.
