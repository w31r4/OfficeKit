## Purpose

This capability gives PPJ a bounded source-bound way to edit one literal bottom
edge of a custom-geometry text rectangle while retaining its representation,
identity, and surrounding topology.

## ADDED Requirements

### Requirement: Project a literal text-rectangle bottom edge

The system SHALL expose one `customGeometryTextRectangleBottomEmu` native leaf
when a recognized custom geometry has exactly one text rectangle whose four
resolved edges are canonical non-negative integers within the shape-local
frame, the rectangle is ordered, and its bottom edge is literal. A direct
numeric `a:rect/@b` token or the exact OfficeKit literal profile
`b="officeKitTextBottom"` with a canonical scaled `a:gd` formula MAY provide
the bottom token. Reference-backed, missing, malformed, extension-bearing,
negative, out-of-frame, or unordered rectangles MUST remain source-owned.

#### Scenario: Project an authored private-profile bottom edge

- **WHEN** an authored/imported custom geometry has a valid text rectangle with
  a bottom edge of `600000` shape-local EMUs represented by the OfficeKit
  `officeKitTextBottom` scaled guide
- **THEN** PPJ exposes one `customGeometryTextRectangleBottomEmu` leaf with
  numeric value `600000`

#### Scenario: Keep a reference-backed bottom edge opaque

- **WHEN** `a:rect/@b` names a built-in or ordinary declared guide instead of a
  direct numeric token or the exact private literal profile
- **THEN** PPJ does not expose a `customGeometryTextRectangleBottomEmu` leaf

### Requirement: Edit only the selected text-rectangle bottom coordinate

The system SHALL accept a source-bound edit only when the expected bottom
value, rectangle structure, all four literal edge tokens, private-guide
identity (if present), and requested in-frame ordered coordinate are valid. A
direct source edit MUST replace only `a:rect/@b`; a private-profile edit MUST
replace only the numeric operand of `officeKitTextBottom`'s scaled
`a:gd/@fmla` while retaining the `a:rect` reference and scale. The edit MUST
preserve left/top/right, guides, paths, handles, connection sites, and package
topology. A stale, out-of-frame, unordered, changed-representation, or
unsupported target MUST fail closed without a partial package.

#### Scenario: Edit and reproject a literal bottom edge

- **WHEN** a PPJ edit changes a projected bottom edge from `600000` to `700000`
  with matching source preconditions and the existing rectangle remains
  ordered
- **THEN** only the owning SlidePart changes, the bottom edge resolves to
  `700000`, and a second projection reports `700000`

#### Scenario: Reject a stale or invalid bottom-edge edit

- **WHEN** the source literal/guide no longer contains the expected bottom value
  or the requested coordinate is negative, beyond the shape height, or not
  greater than the top edge
- **THEN** the edit fails closed with diagnostics and does not write the
  source-bound package

#### Scenario: Preserve text-rectangle representation and topology

- **WHEN** the selected bottom edge is edited
- **THEN** the `officeKitTextBottom` representation (when present), left/top/
  right edges, guide identity, handles, connection sites, paths, and other
  geometry topology remain unchanged
