## Purpose

This capability gives PPJ a bounded source-bound way to edit one literal XY
adjustment-handle x position while retaining the handle's guide identity and
the surrounding custom-geometry topology.

## ADDED Requirements

### Requirement: Project a literal XY handle x position

The system SHALL expose an ordered `customGeometryAdjustmentHandleXEmu` native
leaf for each direct `a:ahXY` adjustment handle whose direct `a:pos/@x` is a
canonical non-negative integer within the shape-local frame. The leaf value
SHALL use shape-local EMU units and its native index SHALL identify the handle
in the direct `a:ahLst` order, including any intervening polar handles.
Formula-backed, missing, malformed, extension-bearing, non-XY, or out-of-frame
x positions MUST remain source-owned and MUST NOT be presented as this leaf.

#### Scenario: Project a direct literal XY handle x position

- **WHEN** an imported source-bound custom geometry has an ordered `a:ahXY`
  handle whose direct position contains `x="180000"` inside the shape frame
- **THEN** PPJ exposes one `customGeometryAdjustmentHandleXEmu` leaf at that
  handle's native index with numeric value `180000`

#### Scenario: Keep a formula-backed handle position opaque

- **WHEN** the selected XY handle's position x references a guide or built-in
  formula instead of a literal integer
- **THEN** PPJ does not expose a
  `customGeometryAdjustmentHandleXEmu` leaf for that handle

### Requirement: Edit only the selected XY handle x token

The system SHALL accept a source-bound edit only when the native handle index,
XY handle identity, direct `ahXY/pos` structure, expected source token, and
requested coordinate are all valid. A successful edit MUST replace only the
selected `a:pos/@x` value in the owning SlidePart, MUST preserve the handle's
guide references, bounds, y position, all other handles and geometry children,
and MUST keep the same package topology. A stale expected value, out-of-range
value, changed handle order/kind, or unsupported XML structure MUST fail closed
without producing a partially edited package.

#### Scenario: Edit and reproject a literal XY handle x position

- **WHEN** a PPJ edit changes a projected literal handle x coordinate from
  `180000` to `240000` with matching source preconditions
- **THEN** only the owning SlidePart changes, the selected `a:pos/@x` becomes
  `240000`, and a second projection reports `240000` at the same native index

#### Scenario: Reject a stale or invalid handle x edit

- **WHEN** the source token no longer contains the expected x value or the
  requested value is negative or outside the shape-local frame
- **THEN** the edit fails closed with diagnostics and does not write the
  source-bound package

#### Scenario: Preserve handle identity and adjacent geometry

- **WHEN** the selected handle x coordinate is edited
- **THEN** its `gdRefX`, `gdRefY`, min/max bounds, y position, handle order and
  kind, guides, paths, connection sites, and other geometry topology remain
  unchanged
