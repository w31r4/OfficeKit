## Purpose

This capability gives PPJ a bounded source-bound way to edit one literal top
edge of a custom-geometry text rectangle while retaining its representation,
identity, and surrounding topology.

## ADDED Requirements

### Requirement: Project a literal text-rectangle top edge

The system SHALL expose one `customGeometryTextRectangleTopEmu` native leaf
when a recognized custom geometry has exactly one text rectangle whose four
resolved edges are canonical non-negative integers within the shape-local
frame, the rectangle is ordered, and its top edge is literal. A direct numeric
`a:rect/@t` token or the exact OfficeKit literal profile
`t="officeKitTextTop"` with a canonical scaled `a:gd` formula MAY provide the
top token. Reference-backed, missing, malformed, extension-bearing, negative,
out-of-frame, or unordered rectangles MUST remain source-owned.

#### Scenario: Project an authored private-profile top edge

- **WHEN** an authored/imported custom geometry has a valid text rectangle with
  a top edge of `200000` shape-local EMUs represented by the OfficeKit
  `officeKitTextTop` scaled guide
- **THEN** PPJ exposes one `customGeometryTextRectangleTopEmu` leaf with
  numeric value `200000`

#### Scenario: Keep a reference-backed top edge opaque

- **WHEN** `a:rect/@t` names a built-in or ordinary declared guide instead of a
  direct numeric token or the exact private literal profile
- **THEN** PPJ does not expose a `customGeometryTextRectangleTopEmu` leaf

### Requirement: Edit only the selected text-rectangle top coordinate

The system SHALL accept a source-bound edit only when the expected top value,
rectangle structure, all four literal edge tokens, private-guide identity (if
present), and requested in-frame ordered coordinate are valid. A direct source
edit MUST replace only `a:rect/@t`; a private-profile edit MUST replace only
the numeric operand of `officeKitTextTop`'s scaled `a:gd/@fmla` while retaining
the `a:rect` reference and scale. The edit MUST preserve left/right/bottom,
guides, paths, handles, connection sites, and package topology. A stale,
out-of-frame, unordered, changed-representation, or unsupported target MUST
fail closed without a partial package.

#### Scenario: Edit and reproject a literal top edge

- **WHEN** a PPJ edit changes a projected top edge from `200000` to `260000`
  with matching source preconditions and the existing rectangle remains
  ordered
- **THEN** only the owning SlidePart changes, the top edge resolves to
  `260000`, and a second projection reports `260000`

#### Scenario: Reject a stale or invalid top-edge edit

- **WHEN** the source literal/guide no longer contains the expected top value or
  the requested coordinate is negative, beyond the shape height, or not less
  than the bottom edge
- **THEN** the edit fails closed with diagnostics and does not write the
  source-bound package

#### Scenario: Preserve text-rectangle representation and topology

- **WHEN** the selected top edge is edited
- **THEN** the `officeKitTextTop` representation (when present),
  left/right/bottom edges, guide identity, handles, connection sites, paths,
  and other geometry topology remain unchanged
