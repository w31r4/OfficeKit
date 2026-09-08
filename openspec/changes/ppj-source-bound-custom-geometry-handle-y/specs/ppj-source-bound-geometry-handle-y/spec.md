## Purpose

This capability gives PPJ a bounded source-bound way to edit one literal XY
adjustment-handle y position while retaining handle identity and custom
geometry topology.

## ADDED Requirements

### Requirement: Project a literal XY handle y position

The system SHALL expose an ordered `customGeometryAdjustmentHandleYEmu` native
leaf for each direct `a:ahXY` adjustment handle whose direct `a:pos/@y` is a
canonical non-negative integer within the shape-local frame. The leaf value
SHALL use shape-local EMU units and its native index SHALL identify the handle
in the direct `a:ahLst` order, including any intervening polar handles.
Formula-backed, missing, malformed, extension-bearing, non-XY, or out-of-frame
y positions MUST remain source-owned and MUST NOT be presented as this leaf.

#### Scenario: Project a direct literal XY handle y position

- **WHEN** an imported source-bound custom geometry has an ordered `a:ahXY`
  handle whose direct position contains `y="240000"` inside the shape frame
- **THEN** PPJ exposes one `customGeometryAdjustmentHandleYEmu` leaf at that
  handle's native index with numeric value `240000`

#### Scenario: Keep a formula-backed handle position opaque

- **WHEN** the selected XY handle's position y references a guide or built-in
  formula instead of a literal integer
- **THEN** PPJ does not expose a
  `customGeometryAdjustmentHandleYEmu` leaf for that handle

### Requirement: Edit only the selected XY handle y token

The system SHALL accept a source-bound edit only when the native handle index,
XY handle identity, direct `ahXY/pos` structure, expected source token, and
requested coordinate are all valid. A successful edit MUST replace only the
selected `a:pos/@y` value in the owning SlidePart, MUST preserve the handle's
guide references, bounds, x position, all other handles and geometry children,
and MUST keep the same package topology. A stale expected value, out-of-range
value, changed handle order/kind, or unsupported XML structure MUST fail closed
without producing a partially edited package.

#### Scenario: Edit and reproject a literal XY handle y position

- **WHEN** a PPJ edit changes a projected literal handle y coordinate from
  `240000` to `360000` with matching source preconditions
- **THEN** only the owning SlidePart changes, the selected `a:pos/@y` becomes
  `360000`, and a second projection reports `360000` at the same native index

#### Scenario: Reject a stale or invalid handle y edit

- **WHEN** the source token no longer contains the expected y value or the
  requested value is negative or outside the shape-local frame
- **THEN** the edit fails closed with diagnostics and does not write the
  source-bound package

#### Scenario: Preserve handle identity and adjacent geometry

- **WHEN** the selected handle y coordinate is edited
- **THEN** its `gdRefX`, `gdRefY`, min/max bounds, x position, handle order and
  kind, guides, paths, connection sites, and other geometry topology remain
  unchanged
