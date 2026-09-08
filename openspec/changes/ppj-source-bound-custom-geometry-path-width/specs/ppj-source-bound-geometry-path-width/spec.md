## Purpose

Provides a narrow, source-faithful PPJ capability for editing explicit custom-
geometry path viewport widths without exposing unsupported path topology.

## ADDED Requirements

### Requirement: Project a literal custom-path width

The system SHALL expose one `customGeometryPathWidth` native leaf for each
recognized custom-geometry path whose direct `a:path/@w` is a positive canonical
integer no greater than `int.MaxValue`. The leaf's `NativeLeafIndex` SHALL be
the path's zero-based order within the direct path list, and its value SHALL be
the DrawingML path-coordinate width. An omitted, zero, negative, out-of-range,
malformed, extension-bearing, or unsupported path SHALL remain source-owned.

#### Scenario: Positive path width is projected

- **WHEN** an imported custom geometry contains a recognized path at index zero
  with direct `a:path/@w="1000"`
- **THEN** the projected native leaves contain `customGeometryPathWidth` at
  index zero with value `1000`

#### Scenario: Omitted path width stays source-owned

- **WHEN** a custom geometry path has no direct `@w` token or has width zero in
  the typed representation
- **THEN** no `customGeometryPathWidth` leaf is emitted for that path

### Requirement: Edit only the selected path width

The system SHALL source-bind a `customGeometryPathWidth` edit to the selected
direct `a:path/@w` token and SHALL change only that token. The requested value
MUST be a changed positive canonical integer no greater than `int.MaxValue`; the
edit MUST preserve `@h`, fill/stroke/extrusion attributes, path commands,
sibling paths, and all non-target package parts. A stale expected width or an
unsupported path structure MUST fail closed.

#### Scenario: Source-bound width edit round-trips

- **WHEN** a projected path-width leaf changes from `1000` to `1200`
- **THEN** only the owning slide part changes, direct `a:path/@w` becomes
  `1200`, all other path content remains intact, and a second projection
  reports value `1200` at the same path index

#### Scenario: Invalid or stale width edit is rejected

- **WHEN** the requested width is zero, non-canonical, out of range, unchanged,
  or the source no longer contains the expected width/path structure
- **THEN** compilation fails without fabricating a replacement path or
  rewriting unrelated geometry
