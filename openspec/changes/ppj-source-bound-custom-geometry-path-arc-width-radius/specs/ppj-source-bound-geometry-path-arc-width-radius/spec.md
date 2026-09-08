# Custom geometry arc width radius

## ADDED Requirements

### Requirement: expose one bounded arc width radius

The system MUST expose a `customGeometryPathArcWidthRadius` native leaf only
for a direct ordered `a:path/a:arcTo` with exactly the four standard arc
attributes, no child content, and a positive canonical literal `wR` within the
bounded DrawingML coordinate range, when the surrounding geometry passes the
existing recognition profile.

#### Scenario: project a literal arc width radius

- **WHEN** an imported custom geometry contains a recognized path with a
  positive literal `a:arcTo/@wR`
- **THEN** the projected shape exposes that radius with path order in
  `nativeLeafIndex` and direct command order in `textLeafIndex`, while `hR`,
  start/sweep angles, and the rest of the path graph remain source-owned

#### Scenario: edit only the selected arc radius token

- **WHEN** a source-bound PPJ edit changes an issued arc width-radius leaf to
  another positive canonical bounded radius
- **THEN** only that arc's `wR` value changes in the owning SlidePart and a
  subsequent projection reports the new radius

#### Scenario: reject reference-backed or unsupported arc state

- **WHEN** the selected radius is formula/reference-backed, missing, malformed,
  noncanonical, extension-bearing, non-positive, out of range, stale, or the
  surrounding geometry cannot pass the bounded profile
- **THEN** the system MUST keep the radius source-owned or fail closed rather
  than apply a partial arc edit
