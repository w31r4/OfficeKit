## Purpose

This capability gives PPJ a bounded source-bound way to edit one literal
custom-geometry connection-site y coordinate while keeping the surrounding
DrawingML geometry opaque and source-faithful.

## ADDED Requirements

### Requirement: Project a literal connection-site y coordinate

The system SHALL expose an ordered
`customGeometryConnectionSiteYEmu` native leaf for each direct custom-geometry
connection site whose `a:pos/@y` is a canonical non-negative signed integer
within the shape-local frame. The leaf value SHALL use shape-local EMU units
and its native index SHALL identify the connection-site position in the
`a:cxnLst` order. Formula-backed, missing, malformed, extension-bearing, or
out-of-frame y coordinates MUST remain source-owned and MUST NOT be presented
as this leaf.

#### Scenario: Project a direct literal y coordinate

- **WHEN** an imported source-bound custom geometry has a direct ordered
  connection site with `a:pos/@y="240000"` inside the shape frame
- **THEN** PPJ exposes one `customGeometryConnectionSiteYEmu` leaf at that
  site's native index with numeric value `240000`

#### Scenario: Keep a formula-backed y coordinate opaque

- **WHEN** the selected connection site's `a:pos/@y` references a guide or
  built-in formula instead of a literal integer
- **THEN** PPJ does not expose a
  `customGeometryConnectionSiteYEmu` leaf for that site

### Requirement: Edit only the selected connection-site y token

The system SHALL accept a source-bound edit only when the native site index,
direct `cxnLst/cxn/pos` structure, expected source token, and requested
coordinate are all valid. A successful edit MUST replace only the selected
`a:pos/@y` value in the owning SlidePart, MUST preserve the selected site's x
and angle values, all other sites and geometry children, and MUST keep the
same package topology. A stale expected value, out-of-range value, changed
site order, or unsupported XML structure MUST fail closed without producing a
partially edited package.

#### Scenario: Edit and reproject a literal y coordinate

- **WHEN** a PPJ edit changes a projected literal y coordinate from `240000`
  to `360000` with matching source preconditions
- **THEN** only the owning SlidePart changes, the selected `a:pos/@y` becomes
  `360000`, and a second projection reports `360000` at the same native index

#### Scenario: Reject a stale or invalid y edit

- **WHEN** the source token no longer contains the expected y value or the
  requested value is negative or outside the shape-local frame
- **THEN** the edit fails closed with diagnostics and does not write the
  source-bound package

#### Scenario: Preserve adjacent connection-site semantics

- **WHEN** the selected y coordinate is edited in a list containing other
  sites, formula-backed angles/positions, guides, handles, or paths
- **THEN** those values, formulas, order, and topology remain unchanged
