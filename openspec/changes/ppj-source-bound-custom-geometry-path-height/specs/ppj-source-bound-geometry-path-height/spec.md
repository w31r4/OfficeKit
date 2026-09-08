## Purpose

Provides a narrow, source-faithful PPJ capability for editing explicit custom-
geometry path viewport heights without exposing unsupported path topology.

## ADDED Requirements

### Requirement: Project a literal custom-path height

The system SHALL expose one `customGeometryPathHeight` native leaf for each
recognized custom-geometry path whose direct `a:path/@h` is a positive canonical
integer no greater than `int.MaxValue`. The leaf's `NativeLeafIndex` SHALL be
the path's zero-based order within the direct path list, and its value SHALL be
the DrawingML path-coordinate height. An omitted, zero, negative, out-of-range,
malformed, extension-bearing, or unsupported path SHALL remain source-owned.

#### Scenario: Positive path height is projected

- **WHEN** an imported custom geometry contains a recognized path at index zero
  with direct `a:path/@h="1000"`
- **THEN** the projected native leaves contain `customGeometryPathHeight` at
  index zero with value `1000`

#### Scenario: Omitted path height stays source-owned

- **WHEN** a custom geometry path has no direct `@h` token or has height zero in
  the typed representation
- **THEN** no `customGeometryPathHeight` leaf is emitted for that path

### Requirement: Edit only the selected path height

The system SHALL source-bind a `customGeometryPathHeight` edit to the selected
direct `a:path/@h` token and SHALL change only that token. The requested value
MUST be a changed positive canonical integer no greater than `int.MaxValue`; the
edit MUST preserve `@w`, fill/stroke/extrusion attributes, path commands,
sibling paths, and all non-target package parts. A stale expected height or an
unsupported path structure MUST fail closed.

#### Scenario: Source-bound height edit round-trips

- **WHEN** a projected path-height leaf changes from `1000` to `1200`
- **THEN** only the owning slide part changes, direct `a:path/@h` becomes
  `1200`, all other path content remains intact, and a second projection
  reports value `1200` at the same path index

#### Scenario: Invalid or stale height edit is rejected

- **WHEN** the requested height is zero, non-canonical, out of range, unchanged,
  or the source no longer contains the expected height/path structure
- **THEN** compilation fails without fabricating a replacement path or
  rewriting unrelated geometry
