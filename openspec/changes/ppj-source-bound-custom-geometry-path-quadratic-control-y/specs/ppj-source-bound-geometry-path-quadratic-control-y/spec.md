# Custom geometry quadratic control-point y coordinate

## ADDED Requirements

### Requirement: expose one bounded quadratic control-point y coordinate

The system MUST expose a `customGeometryPathQuadraticControlY` native leaf
only for a direct ordered `a:path/a:quadBezTo` with exactly two direct `a:pt`
children, where the first point's `y` is a canonical signed integer within the
bounded DrawingML coordinate range and the geometry passes the existing
recognition profile.

#### Scenario: project a literal quadratic control-point coordinate

- **WHEN** an imported custom geometry contains a recognized path with a
  literal quadratic control-point y coordinate
- **THEN** the projected shape exposes that coordinate with path order in
  `nativeLeafIndex` and direct command order in `textLeafIndex`, while the
  control-point x, end point, and rest of the path graph remain source-owned

#### Scenario: edit only the selected control-point y token

- **WHEN** a source-bound PPJ edit changes an issued quadratic control-point y
  leaf to another canonical bounded coordinate
- **THEN** only that first point's `y` value changes in the owning SlidePart and
  a subsequent projection reports the new coordinate

#### Scenario: reject reference-backed or unsupported command state

- **WHEN** the selected y is formula/reference-backed, malformed, noncanonical,
  extension-bearing, out of range, stale, the quadratic point topology is not
  exactly one control point plus one end point, or the surrounding geometry
  cannot pass the bounded profile
- **THEN** the system MUST keep the coordinate source-owned or fail closed
  rather than issue or apply a partial path edit
