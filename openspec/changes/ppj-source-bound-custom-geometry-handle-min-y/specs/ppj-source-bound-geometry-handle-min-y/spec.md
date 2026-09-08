## Purpose

This capability gives PPJ a bounded source-bound way to edit one literal
minimum-y bound on an XY adjustment handle while retaining its identity and
the surrounding custom-geometry topology.

## ADDED Requirements

### Requirement: Project a literal XY handle minimum-y bound

The system SHALL expose an ordered `customGeometryAdjustmentHandleMinYEmu`
native leaf for each direct `a:ahXY` adjustment handle whose `@minY` is a
canonical non-negative integer within the shape-local frame. The leaf value
SHALL use shape-local EMU units and its native index SHALL identify the handle
in the direct `a:ahLst` order, including any intervening polar handles.
Formula-backed, missing, malformed, extension-bearing, non-XY, or out-of-frame
bounds MUST remain source-owned and MUST NOT be presented as this leaf.

#### Scenario: Project a direct literal minimum-y bound

- **WHEN** an imported source-bound custom geometry has an ordered `a:ahXY`
  handle with `minY="20000"` inside the shape frame
- **THEN** PPJ exposes one `customGeometryAdjustmentHandleMinYEmu` leaf at
  that handle's native index with numeric value `20000`

#### Scenario: Keep a reference-backed minimum-y bound opaque

- **WHEN** the selected XY handle's `minY` names a built-in or declared guide
  instead of a literal integer
- **THEN** PPJ does not expose a
  `customGeometryAdjustmentHandleMinYEmu` leaf for that handle

### Requirement: Edit only the selected minimum-y token

The system SHALL accept a source-bound edit only when the native handle index,
XY handle identity, direct handle structure, expected source token, and
requested bound are all valid. A successful edit MUST replace only the
selected `a:ahXY/@minY` value in the owning SlidePart, MUST preserve guide
references, x bounds, `maxY`, handle position, all other handles and geometry
children, and MUST keep the same package topology. A stale expected value,
out-of-frame value, changed handle order/kind, or unsupported XML structure
MUST fail closed without producing a partially edited package.

#### Scenario: Edit and reproject a literal minimum-y bound

- **WHEN** a PPJ edit changes a projected literal minimum-y bound from `20000`
  to `40000` with matching source preconditions and a valid existing handle
  range
- **THEN** only the owning SlidePart changes, the selected `@minY` becomes
  `40000`, and a second projection reports `40000` at the same native index

#### Scenario: Reject a stale or invalid minimum-y edit

- **WHEN** the source token no longer contains the expected bound or the
  requested value is negative or outside the shape-local frame
- **THEN** the edit fails closed with diagnostics and does not write the
  source-bound package

#### Scenario: Preserve handle identity and range topology

- **WHEN** the selected minimum-y bound is edited
- **THEN** `gdRefX/gdRefY`, x bounds, `maxY`, position, handle order and kind,
  guides, paths, connection sites, and other geometry topology remain
  unchanged
