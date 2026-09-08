# Custom geometry cubic end-point y coordinate

## ADDED Requirements

### Requirement: expose one bounded cubic end-point y coordinate

The system MUST expose a `customGeometryPathCubicEndY` native leaf only for a
direct ordered `a:path/a:cubicBezTo` with exactly three direct `a:pt` children,
where the third point's `y` is a canonical signed integer within the bounded
DrawingML coordinate range and the geometry passes the existing recognition
profile.

#### Scenario: project a literal cubic end-point coordinate

- **WHEN** an imported custom geometry contains a recognized path with a
  literal cubic end-point y coordinate
- **THEN** the projected shape exposes that coordinate with path order in
  `nativeLeafIndex` and direct command order in `textLeafIndex`, while both
  control points, the paired end-point x, and the rest of the path graph remain
  source-owned

#### Scenario: edit only the selected end-point y token

- **WHEN** a source-bound PPJ edit changes an issued cubic end-point y leaf to
  another canonical bounded coordinate
- **THEN** only that third point's `y` value changes in the owning SlidePart and
  a subsequent projection reports the new coordinate

#### Scenario: reject reference-backed or unsupported command state

- **WHEN** the selected y is formula/reference-backed, malformed, noncanonical,
  extension-bearing, out of range, stale, the cubic point topology is not
  exactly two control points plus one end point, or the surrounding geometry
  cannot pass the bounded profile
- **THEN** the system MUST keep the coordinate source-owned or fail closed
  rather than issue or apply a partial path edit
