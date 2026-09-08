## Purpose

This capability gives PPJ a bounded source-bound way to edit one literal
custom-geometry connection-site x coordinate while keeping the surrounding
DrawingML geometry opaque and source-faithful.

## ADDED Requirements

### Requirement: Project a literal connection-site x coordinate

The system SHALL expose an ordered
`customGeometryConnectionSiteXEmu` native leaf for each direct custom-geometry
connection site whose `a:pos/@x` is a canonical non-negative signed integer
within the shape-local frame. The leaf value SHALL use shape-local EMU units
and its native index SHALL identify the connection-site position in the
`a:cxnLst` order. Formula-backed, missing, malformed, extension-bearing, or
out-of-frame x coordinates MUST remain source-owned and MUST NOT be presented
as this leaf.

#### Scenario: Project a direct literal x coordinate

- **WHEN** an imported source-bound custom geometry has a direct ordered
  connection site with `a:pos/@x="120000"` inside the shape frame
- **THEN** PPJ exposes one `customGeometryConnectionSiteXEmu` leaf at that
  site's native index with numeric value `120000`

#### Scenario: Keep a formula-backed x coordinate opaque

- **WHEN** the selected connection site's `a:pos/@x` references a guide or
  built-in formula instead of a literal integer
- **THEN** PPJ does not expose a
  `customGeometryConnectionSiteXEmu` leaf for that site

### Requirement: Edit only the selected connection-site x token

The system SHALL accept a source-bound edit only when the native site index,
direct `cxnLst/cxn/pos` structure, expected source token, and requested
coordinate are all valid. A successful edit MUST replace only the selected
`a:pos/@x` value in the owning SlidePart, MUST preserve the selected site's y
and angle values, all other sites and geometry children, and MUST keep the
same package topology. A stale expected value, out-of-range value, changed
site order, or unsupported XML structure MUST fail closed without producing a
partially edited package.

#### Scenario: Edit and reproject a literal x coordinate

- **WHEN** a PPJ edit changes a projected literal x coordinate from `120000`
  to `180000` with matching source preconditions
- **THEN** only the owning SlidePart changes, the selected `a:pos/@x` becomes
  `180000`, and a second projection reports `180000` at the same native index

#### Scenario: Reject a stale or invalid x edit

- **WHEN** the source token no longer contains the expected x value or the
  requested value is negative or outside the shape-local frame
- **THEN** the edit fails closed with diagnostics and does not write the
  source-bound package

#### Scenario: Preserve adjacent connection-site semantics

- **WHEN** the selected x coordinate is edited in a list containing other
  sites, formula-backed angles/positions, guides, handles, or paths
- **THEN** those values, formulas, order, and topology remain unchanged
