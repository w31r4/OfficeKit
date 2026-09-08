## Purpose

This capability gives PPJ a bounded source-bound way to edit one literal
maximum-x bound on an XY adjustment handle while retaining its identity and
the surrounding custom-geometry topology.

## ADDED Requirements

### Requirement: Project a literal XY handle maximum-x bound

The system SHALL expose an ordered `customGeometryAdjustmentHandleMaxXEmu`
native leaf for each direct `a:ahXY` adjustment handle whose `@maxX` is a
canonical non-negative integer within the shape-local frame. The leaf value
SHALL use shape-local EMU units and its native index SHALL identify the handle
in the direct `a:ahLst` order, including any intervening polar handles.
Formula-backed, missing, malformed, extension-bearing, non-XY, or out-of-frame
bounds MUST remain source-owned and MUST NOT be presented as this leaf.

#### Scenario: Project a direct literal maximum-x bound

- **WHEN** an imported source-bound custom geometry has an ordered `a:ahXY`
  handle with `maxX="380000"` inside the shape frame
- **THEN** PPJ exposes one `customGeometryAdjustmentHandleMaxXEmu` leaf at
  that handle's native index with numeric value `380000`

#### Scenario: Keep a reference-backed maximum-x bound opaque

- **WHEN** the selected XY handle's `maxX` names a built-in or declared guide
  instead of a literal integer
- **THEN** PPJ does not expose a
  `customGeometryAdjustmentHandleMaxXEmu` leaf for that handle

### Requirement: Edit only the selected maximum-x token

The system SHALL accept a source-bound edit only when the native handle index,
XY handle identity, direct handle structure, expected source token, and
requested bound are all valid. A successful edit MUST replace only the
selected `a:ahXY/@maxX` value in the owning SlidePart, MUST preserve guide
references, `minX`, y bounds, handle position, all other handles and geometry
children, and MUST keep the same package topology. A stale expected value,
out-of-frame value, changed handle order/kind, or unsupported XML structure
MUST fail closed without producing a partially edited package.

#### Scenario: Edit and reproject a literal maximum-x bound

- **WHEN** a PPJ edit changes a projected literal maximum-x bound from
  `380000` to `420000` with matching source preconditions and a valid existing
  handle range
- **THEN** only the owning SlidePart changes, the selected `@maxX` becomes
  `420000`, and a second projection reports `420000` at the same native index

#### Scenario: Reject a stale or invalid maximum-x edit

- **WHEN** the source token no longer contains the expected bound or the
  requested value is negative or outside the shape-local frame
- **THEN** the edit fails closed with diagnostics and does not write the
  source-bound package

#### Scenario: Preserve handle identity and range topology

- **WHEN** the selected maximum-x bound is edited
- **THEN** `gdRefX/gdRefY`, `minX`, y bounds, position, handle order and kind,
  guides, paths, connection sites, and other geometry topology remain
  unchanged

