# Custom geometry move-to y coordinate

## ADDED Requirements

### Requirement: expose one bounded move-to y coordinate

The system MUST expose a `customGeometryPathMoveToY` native leaf only for a
direct ordered `a:path/a:moveTo/a:pt` whose `y` is a canonical signed integer
within the bounded DrawingML coordinate range and whose geometry passes the
existing recognition profile.

#### Scenario: project a literal move-to coordinate

- **WHEN** an imported custom geometry contains a recognized path with a
  literal move-to point y coordinate
- **THEN** the projected shape exposes that coordinate with path order in
  `nativeLeafIndex` and direct command order in `textLeafIndex`, while the
  paired x and the rest of the path graph remain source-owned

#### Scenario: edit only the selected y token

- **WHEN** a source-bound PPJ edit changes an issued move-to y leaf to another
  canonical bounded coordinate
- **THEN** only that point's `y` value changes in the owning SlidePart and a
  subsequent projection reports the new coordinate

#### Scenario: reject reference-backed or unsupported command state

- **WHEN** the selected y is formula/reference-backed, malformed, noncanonical,
  extension-bearing, out of range, stale, or the surrounding geometry cannot
  pass the bounded profile
- **THEN** the system MUST keep the coordinate source-owned or fail closed
  rather than issue or apply a partial path edit
