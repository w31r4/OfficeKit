## Purpose

This capability gives PPJ a bounded source-bound way to edit one literal
minimum-radius bound on a polar adjustment handle while retaining its identity
and the surrounding custom-geometry topology.

## ADDED Requirements

### Requirement: Project a literal polar-handle minimum radius

The system SHALL expose an ordered
`customGeometryAdjustmentHandlePolarMinRadiusEmu` native leaf for each direct
`a:ahPolar` adjustment handle whose `@minR` is a canonical non-negative
integer within the bounded DrawingML integer range. The leaf value SHALL use
shape-local EMU units and its native index SHALL identify the handle in the
direct `a:ahLst` order, including any intervening XY handles. Formula-backed,
missing, malformed, extension-bearing, non-polar, or out-of-range bounds MUST
remain source-owned and MUST NOT be presented as this leaf.

#### Scenario: Project a direct literal polar minimum radius

- **WHEN** an imported source-bound custom geometry has an ordered `a:ahPolar`
  handle with `minR="100000"` and a valid existing radial range
- **THEN** PPJ exposes one `customGeometryAdjustmentHandlePolarMinRadiusEmu`
  leaf at that handle's native index with numeric value `100000`

#### Scenario: Keep a reference-backed polar minimum radius opaque

- **WHEN** the selected polar handle's `minR` names a built-in or declared
  guide instead of a literal integer
- **THEN** PPJ does not expose a
  `customGeometryAdjustmentHandlePolarMinRadiusEmu` leaf for that handle

### Requirement: Edit only the selected polar minimum-radius token

The system SHALL accept a source-bound edit only when the native handle index,
polar handle identity, direct handle structure, expected source token, and
requested bound are all valid. A successful edit MUST replace only the
selected `a:ahPolar/@minR` value in the owning SlidePart, MUST preserve guide
references, `maxR`, angle bounds, handle position, all other handles and
geometry children, and MUST keep the same package topology. A stale expected
value, negative value, changed handle order/kind, or unsupported XML structure
MUST fail closed without producing a partially edited package.

#### Scenario: Edit and reproject a literal polar minimum radius

- **WHEN** a PPJ edit changes a projected literal polar minimum radius from
  `100000` to `150000` with matching source preconditions and a valid existing
  handle range
- **THEN** only the owning SlidePart changes, the selected `@minR` becomes
  `150000`, and a second projection reports `150000` at the same native index

#### Scenario: Reject a stale or invalid polar minimum-radius edit

- **WHEN** the source token no longer contains the expected bound or the
  requested value is negative or outside the bounded integer range
- **THEN** the edit fails closed with diagnostics and does not write the
  source-bound package

#### Scenario: Preserve polar identity and range topology

- **WHEN** the selected polar minimum-radius bound is edited
- **THEN** `gdRefR/gdRefAng`, `maxR`, angle bounds, position, handle order and
  kind, guides, paths, connection sites, and other geometry topology remain
  unchanged

