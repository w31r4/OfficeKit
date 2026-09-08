# Custom geometry line-to x coordinate

## ADDED Requirements

### Requirement: expose one bounded line-to x coordinate

The system MUST expose a `customGeometryPathLineToX` native leaf only for a
direct ordered `a:path/a:lnTo/a:pt` whose `x` is a canonical signed integer
within the bounded DrawingML coordinate range and whose geometry passes the
existing recognition profile.

#### Scenario: project a literal line-to coordinate

- **WHEN** an imported custom geometry contains a recognized path with a
  literal line-to point x coordinate
- **THEN** the projected shape exposes that coordinate with path order in
  `nativeLeafIndex` and direct command order in `textLeafIndex`, while the
  paired y and the rest of the path graph remain source-owned

#### Scenario: edit only the selected x token

- **WHEN** a source-bound PPJ edit changes an issued line-to x leaf to another
  canonical bounded coordinate
- **THEN** only that point's `x` value changes in the owning SlidePart and a
  subsequent projection reports the new coordinate

#### Scenario: reject reference-backed or unsupported command state

- **WHEN** the selected x is formula/reference-backed, malformed, noncanonical,
  extension-bearing, out of range, stale, or the surrounding geometry cannot
  pass the bounded profile
- **THEN** the system MUST keep the coordinate source-owned or fail closed
  rather than issue or apply a partial path edit
