## Purpose

Expose a direct rich-text run outer-shadow theme color as a typed PPJ field
while preserving the source package and rejecting unsupported effect topology.

## ADDED Requirements

### Requirement: Direct rich-text shadow theme color is projected as a bounded native field

The PPJ projection MUST expose `run.style.shadow.color` as
`textShadowColorScheme` when the run owns exactly one direct outer shadow with
one direct `schemeClr/@val` token accepted by the existing theme-token grammar,
at least one supported bounded geometry attribute, no shadow transform
attributes, no sibling effects, and no unknown descendants. The native leaf
MUST retain the canonical theme token and the typed value MUST be that token.

#### Scenario: Canonical direct outer shadow exposes a theme color

- **WHEN** an imported text run contains one direct `a:effectLst/a:outerShdw`
  with bounded blur, distance, or direction and one direct `schemeClr`
- **THEN** the projected run contains a `nativeRef.leaves[]` item with kind
  `textShadowColorScheme` and the same theme token, and
  `run.style.shadow.color` is represented by that token

#### Scenario: Unsupported theme-color topology stays source-owned

- **WHEN** the run has a missing or malformed scheme token, a transformed color,
  an RGB color, missing supported geometry, a shadow transform, a sibling
  effect, or an unknown descendant
- **THEN** the projection emits no `textShadowColorScheme` leaf for that run

### Requirement: Source-bound theme color edits splice only the owning token

A source-bound edit of `textShadowColorScheme` MUST require a changed canonical
theme token, verify the projected leaf's source proof, and replace only the
existing direct `a:outerShdw/a:schemeClr/@val` token in the owning SlidePart. It
MUST preserve alpha, geometry, topology, and every unrelated OPC part.

#### Scenario: Theme color edit preserves the package footprint

- **WHEN** a projected `textShadowColorScheme` leaf changes from one valid
  token to another
- **THEN** only the owning SlidePart is reported changed, only the scheme
  token changes, and all other shadow XML and package parts remain unchanged

#### Scenario: Invalid or stale edits fail closed

- **WHEN** an edit supplies a malformed or unchanged token, or the source token
  no longer matches the expected proof
- **THEN** the operation fails without rewriting the source package

### Requirement: Edited theme color reprojects to the typed PPJ field

After a successful source-bound edit, a subsequent PPJ projection MUST recover
the new `run.style.shadow.color` theme token and the corresponding
`textShadowColorScheme` native leaf without requiring host-specific rendering.

#### Scenario: Edited theme color survives a second projection

- **WHEN** the edited PPTX is projected again
- **THEN** the run's typed shadow color and native leaf both contain the edited
  theme token
