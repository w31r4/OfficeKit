## Purpose

This capability gives imported rich-text runs one precise outer-shadow blur field in PPJ, so a bounded source-preserving edit can change the native blur token and recover the value after reprojection.

## ADDED Requirements

### Requirement: Direct rich-text shadow blur is a typed native leaf

The presentation projection MUST expose `run.style.shadow.blur` as `textShadowBlurRadiusEmu` only when the run owns exactly one direct `a:outerShdw` with an existing canonical non-negative `blurRad`, one direct RGB or theme color child, no transform attributes, no extra effect-list children, and no unsupported descendants. The leaf value MUST be the native EMU integer, and authored and imported PPJ MUST preserve the same semantic blur in points after the normal PPJ conversion.

#### Scenario: Authored shadow blur survives projection
- **WHEN** an authored rich-text run contains a valid outer shadow with a blur expressed in points
- **THEN** the compiled package and a second PPJ projection contain the same `run.style.shadow.blur`, and the native leaf reports the corresponding `textShadowBlurRadiusEmu` value

#### Scenario: Source-bound edit changes only the blur token
- **WHEN** a projected `textShadowBlurRadiusEmu` leaf is changed to another valid EMU value and applied to its source package
- **THEN** only the owning slide XML changes, the existing outer-shadow color and other attributes remain byte-equivalent, Open XML remains valid, and a second projection reports the new blur

#### Scenario: Unsupported shadow topology stays source-owned
- **WHEN** the run effect list has a missing blur token, a shadow transform, a sibling effect, an extra child, or malformed/invalid outer-shadow content
- **THEN** the projection MUST omit `textShadowBlurRadiusEmu` and source-bound editing MUST fail closed without flattening the run effect
