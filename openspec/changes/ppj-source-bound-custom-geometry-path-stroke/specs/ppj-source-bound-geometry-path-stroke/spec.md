## Purpose

This capability exposes one explicit custom-geometry path stroke token as a
source-bound PPJ boolean while preserving the surrounding DrawingML topology.

## ADDED Requirements

### Requirement: Project explicit custom-path stroke leaves

The system MUST project one `customGeometryPathStroke` native leaf for each
recognized custom-geometry path that has an explicit canonical `stroke` token.
The leaf MUST use the direct path's zero-based order as `NativeLeafIndex` and
map `stroke="1"` to `true` and `stroke="0"` to `false`.

#### Scenario: Project an explicit stroked path

- **WHEN** a source-bound custom geometry contains an ordered path with
  `stroke="1"`
- **THEN** the projected PPJ contains a `customGeometryPathStroke` leaf for
  that path index with boolean value `true`

#### Scenario: Preserve an omitted stroke token

- **WHEN** a custom-geometry path omits `stroke`
- **THEN** the system MUST NOT project a stroke leaf for that path

### Requirement: Source-bound edit changes only the selected stroke token

The system MUST accept a canonical boolean edit for a projected
`customGeometryPathStroke` leaf only when the source element and path index
still prove the selected direct path. The edit MUST change only that path's
`stroke` attribute in its owning SlidePart and MUST preserve viewport extents,
fill/extrusion attributes, commands, sibling paths, and unrelated package
parts. A stale raw token, malformed boolean, unsupported path structure, or
unchanged value MUST fail closed without mutation.

#### Scenario: Edit one path and preserve opaque topology

- **WHEN** a projected path leaf changes from `true` to `false` and the source
  element hash and raw value still match
- **THEN** the owning SlidePart changes its selected direct path to
  `stroke="0"`, while the path's other attributes, commands, sibling paths,
  and unrelated package parts remain unchanged

#### Scenario: Reject stale or invalid edits

- **WHEN** the source stroke token is stale, the requested boolean is
  non-canonical, the path structure is unsupported, or the requested value is
  unchanged
- **THEN** compilation MUST fail closed or produce no mutation, and MUST NOT
  rewrite the selected path or materialize an omitted attribute

### Requirement: Reprojection reports the edited stroke

After a successful source-bound edit, reprojection MUST report the new boolean
value for the same path index and MUST continue to omit the leaf when the
source path has no explicit stroke token.

#### Scenario: Reproject the edited path

- **WHEN** the edited SlidePart is projected again
- **THEN** the same indexed `customGeometryPathStroke` leaf reports the new
  boolean value

#### Scenario: Reproject a path without stroke

- **WHEN** a source-bound path has no explicit `stroke` attribute
- **THEN** reprojection contains no `customGeometryPathStroke` leaf for that
  path
