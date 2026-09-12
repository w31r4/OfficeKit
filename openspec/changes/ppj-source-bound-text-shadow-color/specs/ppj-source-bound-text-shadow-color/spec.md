## Purpose

Expose a bounded imported rich-text outer-shadow RGB color as a complete PPJ field so callers can change that one value while the source presentation remains structurally and byte-wise preserved elsewhere.

## ADDED Requirements

### Requirement: Direct rich-text shadow RGB color is projected as a bounded native field

The PPJ projection MUST expose `run.style.shadow.color` as `textShadowColorRgb` when the run owns exactly one direct outer shadow with one direct RGB `srgbClr/@val` token containing six hexadecimal characters, at least one supported bounded geometry attribute, no shadow transform attributes, no sibling effects, and no unknown descendants. The native leaf MUST retain the canonical six-character color token, and its typed value MUST be the corresponding `#RRGGBB` color.

#### Scenario: Canonical direct outer shadow exposes RGB color

- **WHEN** an imported text run contains one direct `a:effectLst/a:outerShdw` with a bounded blur, distance, or direction token and one direct RGB color child
- **THEN** the projected run contains `nativeRef.leaves[]` with `kind` `textShadowColorRgb`, the native six-hex value, and `run.style.shadow.color` with the same RGB color

#### Scenario: Unsupported shadow topology stays source-owned

- **WHEN** the run contains a theme color, missing or malformed RGB color, missing supported geometry, a transform attribute, a sibling effect, an unknown descendant, or an out-of-range geometry token
- **THEN** the projection omits `textShadowColorRgb` and keeps the source-owned effect graph opaque

### Requirement: Source-bound RGB color edits splice only the owning token

A source-bound edit of `textShadowColorRgb` MUST require a changed canonical six-hex RGB value, verify the projected leaf's source proof, and replace only the existing direct `a:outerShdw/a:srgbClr/@val` token in the owning SlidePart. It MUST preserve the alpha child, all other shadow attributes and children, run and paragraph topology, and every unrelated OPC part.

#### Scenario: RGB edit preserves the package footprint

- **WHEN** a projected `textShadowColorRgb` leaf is changed from one valid RGB value to another
- **THEN** only the owning SlidePart is reported changed and its `srgbClr/@val` contains the requested token while blur, distance, direction, alpha, and other package parts remain unchanged

#### Scenario: Invalid or stale edits fail closed

- **WHEN** an edit supplies a malformed or unchanged RGB value, or its expected source token no longer matches
- **THEN** the operation fails without rewriting the source package

### Requirement: Edited RGB color reprojects to the typed PPJ field

After a successful source-bound edit, a subsequent PPJ projection MUST recover the new `run.style.shadow.color` and the corresponding `textShadowColorRgb` native leaf without requiring host-specific rendering.

#### Scenario: Edited RGB color survives a second projection

- **WHEN** the edited PPTX is projected again
- **THEN** the run's typed shadow color and native leaf both contain the edited RGB value
