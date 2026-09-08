# Custom geometry arc sweep angle

## ADDED Requirements

### Requirement: expose one bounded arc sweep angle

The system MUST expose a `customGeometryPathArcSweepAngle60000` native leaf
only for a direct ordered `a:path/a:arcTo` with exactly the four standard arc
attributes, no child content, and a non-zero canonical literal `swAng` within
the signed one-turn range, when the surrounding geometry passes the existing
recognition profile.

#### Scenario: project a literal arc sweep angle

- **WHEN** an imported custom geometry contains a recognized path with a
  canonical non-zero literal `a:arcTo/@swAng` in the bounded range
- **THEN** the projected shape exposes that angle with path order in
  `nativeLeafIndex` and direct command order in `textLeafIndex`, while both
  radii, the start angle, and the rest of the path graph remain source-owned

#### Scenario: edit only the selected sweep token

- **WHEN** a source-bound PPJ edit changes an issued arc sweep-angle leaf to
  another canonical non-zero bounded angle
- **THEN** only that arc's `swAng` value changes in the owning SlidePart and a
  subsequent projection reports the new sweep

#### Scenario: reject reference-backed or unsupported arc state

- **WHEN** the selected angle is formula/reference-backed, missing, malformed,
  noncanonical, extension-bearing, zero, outside the signed one-turn range,
  stale, or the surrounding geometry cannot pass the bounded profile
- **THEN** the system MUST keep the angle source-owned or fail closed rather
  than apply a partial arc edit
