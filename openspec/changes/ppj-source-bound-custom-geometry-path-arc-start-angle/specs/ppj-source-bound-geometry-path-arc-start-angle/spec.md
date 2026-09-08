# Custom geometry arc start angle

## ADDED Requirements

### Requirement: expose one bounded arc start angle

The system MUST expose a `customGeometryPathArcStartAngle60000` native leaf
only for a direct ordered `a:path/a:arcTo` with exactly the four standard arc
attributes, no child content, and a canonical literal `stAng` within the
signed one-turn range, when the surrounding geometry passes the existing
recognition profile.

#### Scenario: project a literal arc start angle

- **WHEN** an imported custom geometry contains a recognized path with a
  canonical literal `a:arcTo/@stAng` in the bounded range
- **THEN** the projected shape exposes that angle with path order in
  `nativeLeafIndex` and direct command order in `textLeafIndex`, while both
  radii, the sweep angle, and the rest of the path graph remain source-owned

#### Scenario: edit only the selected angle token

- **WHEN** a source-bound PPJ edit changes an issued arc start-angle leaf to
  another canonical bounded angle
- **THEN** only that arc's `stAng` value changes in the owning SlidePart and a
  subsequent projection reports the new angle

#### Scenario: reject reference-backed or unsupported arc state

- **WHEN** the selected angle is formula/reference-backed, missing, malformed,
  noncanonical, extension-bearing, outside the signed one-turn range, stale,
  or the surrounding geometry cannot pass the bounded profile
- **THEN** the system MUST keep the angle source-owned or fail closed rather
  than apply a partial arc edit
