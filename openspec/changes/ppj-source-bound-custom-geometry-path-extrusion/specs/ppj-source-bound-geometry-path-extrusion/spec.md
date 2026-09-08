## Purpose

This capability exposes one explicit custom-geometry path extrusion permission
as a source-bound PPJ boolean while preserving the surrounding DrawingML path.

## ADDED Requirements

### Requirement: Project explicit custom-path extrusion permission

The system MUST project one `customGeometryPathExtrusionAllowed` native leaf
for each recognized custom-geometry path that has an explicit canonical
`extrusionOk` boolean token. The leaf MUST use the direct path's zero-based
order as `NativeLeafIndex` and preserve the token's boolean value.

#### Scenario: Project an explicitly permitted path

- **WHEN** a source-bound custom geometry contains an ordered path with
  `extrusionOk="1"`
- **THEN** the projected PPJ contains a
  `customGeometryPathExtrusionAllowed` leaf for that path index with boolean
  value `true`

#### Scenario: Preserve an omitted extrusion token

- **WHEN** a custom-geometry path omits `extrusionOk`
- **THEN** the system MUST NOT project an extrusion-permission leaf for that
  path

### Requirement: Source-bound edit changes only the selected extrusion token

The system MUST accept a canonical boolean edit for a projected
`customGeometryPathExtrusionAllowed` leaf only when the source element and path
index still prove the selected direct path. The edit MUST change only that
path's `extrusionOk` attribute in its owning SlidePart and MUST preserve
viewport extents, fill/stroke attributes, commands, sibling paths, and
unrelated package parts. A stale raw token, malformed boolean, unsupported path
structure, or unchanged value MUST fail closed or produce no mutation.

#### Scenario: Edit one path and preserve opaque topology

- **WHEN** a projected path leaf changes from `true` to `false` and the source
  element hash and raw value still match
- **THEN** the owning SlidePart changes its selected direct path to
  `extrusionOk="0"`, while the path's other attributes, commands, sibling
  paths, and unrelated package parts remain unchanged

#### Scenario: Reject stale or invalid edits

- **WHEN** the source extrusion token is stale, the requested boolean is
  non-canonical, the path structure is unsupported, or the requested value is
  unchanged
- **THEN** compilation MUST fail closed or produce no mutation, and MUST NOT
  rewrite the selected path or materialize an omitted attribute

### Requirement: Reprojection reports the edited extrusion permission

After a successful source-bound edit, reprojection MUST report the new boolean
value for the same path index and MUST continue to omit the leaf when the
source path has no explicit extrusion token.

#### Scenario: Reproject the edited path

- **WHEN** the edited SlidePart is projected again
- **THEN** the same indexed `customGeometryPathExtrusionAllowed` leaf reports
  the new boolean value

#### Scenario: Reproject a path without extrusion permission

- **WHEN** a source-bound path has no explicit `extrusionOk` attribute
- **THEN** reprojection contains no `customGeometryPathExtrusionAllowed` leaf
  for that path
