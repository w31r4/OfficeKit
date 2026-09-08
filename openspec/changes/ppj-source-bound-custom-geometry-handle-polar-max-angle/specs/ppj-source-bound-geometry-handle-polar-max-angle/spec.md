## Purpose

This capability gives PPJ a bounded source-bound way to edit one literal
maximum-angle bound on a polar adjustment handle while retaining its identity
and the surrounding custom-geometry topology.

## ADDED Requirements

### Requirement: Project a literal polar-handle maximum angle

The system SHALL expose an ordered
`customGeometryAdjustmentHandlePolarMaxAngle60000` native leaf for each direct
`a:ahPolar` adjustment handle whose `@maxAng` is a canonical integer in the
inclusive range -21600000 through 21600000. The leaf value SHALL use
1/60000-degree units and its native index SHALL identify the handle in the
direct `a:ahLst` order, including any intervening XY handles. Formula-backed,
missing, malformed, extension-bearing, non-polar, or out-of-range bounds MUST
remain source-owned and MUST NOT be presented as this leaf.

#### Scenario: Project a direct literal polar maximum angle

- **WHEN** an imported source-bound custom geometry has an ordered `a:ahPolar`
  handle with `maxAng="1800000"` and a valid existing angular range
- **THEN** PPJ exposes one `customGeometryAdjustmentHandlePolarMaxAngle60000`
  leaf at that handle's native index with numeric value `1800000`

#### Scenario: Keep a reference-backed polar maximum angle opaque

- **WHEN** the selected polar handle's `maxAng` names a built-in or declared
  guide instead of a literal integer
- **THEN** PPJ does not expose a
  `customGeometryAdjustmentHandlePolarMaxAngle60000` leaf for that handle

### Requirement: Edit only the selected polar maximum-angle token

The system SHALL accept a source-bound edit only when the native handle index,
polar handle identity, direct handle structure, expected source token, and
requested bound are all valid. A successful edit MUST replace only the selected
`a:ahPolar/@maxAng` value in the owning SlidePart, MUST preserve guide
references, `minAng`, radial bounds, handle position, all other handles and
geometry children, and MUST keep the same package topology. A stale expected
value, out-of-range value, changed handle order/kind, or unsupported XML
structure MUST fail closed without producing a partially edited package.

#### Scenario: Edit and reproject a literal polar maximum angle

- **WHEN** a PPJ edit changes a projected literal polar maximum angle from
  `1800000` to `1200000` with matching source preconditions and a valid
  existing handle range
- **THEN** only the owning SlidePart changes, the selected `@maxAng` becomes
  `1200000`, and a second projection reports `1200000` at the same native index

#### Scenario: Reject a stale or invalid polar maximum-angle edit

- **WHEN** the source token no longer contains the expected bound or the
  requested value is outside the inclusive one-turn angle range
- **THEN** the edit fails closed with diagnostics and does not write the
  source-bound package

#### Scenario: Preserve polar identity and range topology

- **WHEN** the selected polar maximum-angle bound is edited
- **THEN** `gdRefR/gdRefAng`, radial bounds, `minAng`, position, handle order and
  kind, guides, paths, connection sites, and other geometry topology remain
  unchanged
