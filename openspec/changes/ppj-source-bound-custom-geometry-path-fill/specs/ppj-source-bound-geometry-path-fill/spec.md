## Purpose

Provides a narrow, source-faithful PPJ capability for editing explicit custom-
geometry path fill modes without exposing unsupported path topology.

## ADDED Requirements

### Requirement: Project an explicit custom-path fill mode

The system SHALL expose one `customGeometryPathFill` native leaf for each
recognized custom-geometry path whose direct `a:path/@fill` is `norm` or `none`.
The leaf's `NativeLeafIndex` SHALL be the path's zero-based order within the
direct path list, and its value SHALL be `true` for `norm` and `false` for
`none`. An omitted, unsupported, malformed, extension-bearing, or unknown path
fill state SHALL remain source-owned.

#### Scenario: Explicit normal fill is projected

- **WHEN** an imported custom geometry contains a recognized path at index zero
  with direct `a:path/@fill="norm"`
- **THEN** the projected native leaves contain `customGeometryPathFill` at
  index zero with boolean value `true`

#### Scenario: Omitted fill stays source-owned

- **WHEN** a custom geometry path has no direct `@fill` token
- **THEN** no `customGeometryPathFill` leaf is emitted for that path

### Requirement: Edit only the selected path fill token

The system SHALL source-bind a `customGeometryPathFill` edit to the selected
direct `a:path/@fill` token and SHALL change only that token, mapping `true` to
`norm` and `false` to `none`. The edit MUST preserve viewport attributes,
stroke/extrusion attributes, path commands, sibling paths, and all non-target
package parts. A stale raw token or unsupported path structure MUST fail
closed; an unchanged value MUST produce no mutation.

#### Scenario: Source-bound fill edit round-trips

- **WHEN** a projected path-fill leaf changes from `true` to `false`
- **THEN** only the owning slide part changes, direct `a:path/@fill` becomes
  `none`, all other path content remains intact, and a second projection reports
  boolean value `false` at the same path index

#### Scenario: Invalid or stale fill edit is rejected

- **WHEN** the requested value is non-canonical, or the source no longer
  contains the expected fill token/path structure
- **THEN** compilation fails without fabricating a fill attribute or rewriting
  unrelated geometry
