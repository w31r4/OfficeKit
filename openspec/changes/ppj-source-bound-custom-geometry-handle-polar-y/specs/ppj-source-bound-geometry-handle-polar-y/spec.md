## Purpose

This capability gives PPJ a bounded source-bound way to edit one literal y
coordinate on a polar adjustment handle while retaining its identity and the
surrounding custom-geometry topology.

## ADDED Requirements

### Requirement: Project a literal polar-handle position y

The system SHALL expose an ordered
`customGeometryAdjustmentHandlePolarYEmu` native leaf for each direct
`a:ahPolar/a:pos` whose `@y` is a canonical non-negative integer no greater
than the shape-local custom-geometry height. The leaf value SHALL use
shape-local EMU units and its native index SHALL identify the handle in the
direct `a:ahLst` order, including any intervening XY handles. Formula-backed,
missing, malformed, extension-bearing, non-polar, or out-of-frame coordinates
MUST remain source-owned and MUST NOT be presented as this leaf.

#### Scenario: Project a direct literal polar position y

- **WHEN** an imported source-bound custom geometry has an ordered `a:ahPolar`
  handle whose direct position has `y="420000"` within the shape-local frame
- **THEN** PPJ exposes one `customGeometryAdjustmentHandlePolarYEmu` leaf at
  that handle's native index with numeric value `420000`

#### Scenario: Keep a reference-backed polar position y opaque

- **WHEN** the selected polar position's `y` names a built-in or declared guide
  instead of a literal coordinate
- **THEN** PPJ does not expose a
  `customGeometryAdjustmentHandlePolarYEmu` leaf for that handle

### Requirement: Edit only the selected polar position y token

The system SHALL accept a source-bound edit only when the native handle index,
polar handle identity, direct handle structure, expected source token, and
requested coordinate are all valid. A successful edit MUST replace only the
selected `a:ahPolar/a:pos/@y` value in the owning SlidePart, MUST preserve
guide references, range bounds, x, all other handles and geometry children, and
MUST keep the same package topology. A stale expected value, out-of-frame
value, changed handle order/kind, or unsupported XML structure MUST fail closed
without producing a partially edited package.

#### Scenario: Edit and reproject a literal polar position y

- **WHEN** a PPJ edit changes a projected literal polar y from `420000` to
  `520000` with matching source preconditions and an in-frame coordinate
- **THEN** only the owning SlidePart changes, the selected `@y` becomes
  `520000`, and a second projection reports `520000` at the same native index

#### Scenario: Reject a stale or invalid polar position y edit

- **WHEN** the source token no longer contains the expected coordinate or the
  requested value is negative or beyond the shape-local height
- **THEN** the edit fails closed with diagnostics and does not write the
  source-bound package

#### Scenario: Preserve polar identity and topology

- **WHEN** the selected polar position y is edited
- **THEN** `gdRefR/gdRefAng`, radial/angular bounds, x, handle order and kind,
  guides, paths, connection sites, and other geometry topology remain
  unchanged
