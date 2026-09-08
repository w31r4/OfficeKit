## Purpose

This capability gives PPJ a bounded source-bound way to edit one literal right
edge of a custom-geometry text rectangle while retaining its representation,
identity, and surrounding topology.

## ADDED Requirements

### Requirement: Project a literal text-rectangle right edge

The system SHALL expose one `customGeometryTextRectangleRightEmu` native leaf
when a recognized custom geometry has exactly one text rectangle whose four
resolved edges are canonical non-negative integers within the shape-local
frame, the rectangle is ordered, and its right edge is literal. A direct
numeric `a:rect/@r` token or the exact OfficeKit literal profile
`r="officeKitTextRight"` with a canonical scaled `a:gd` formula MAY provide the
right token. Reference-backed, missing, malformed, extension-bearing,
negative, out-of-frame, or unordered rectangles MUST remain source-owned.

#### Scenario: Project an authored private-profile right edge

- **WHEN** an authored/imported custom geometry has a valid text rectangle with
  a right edge of `800000` shape-local EMUs represented by the OfficeKit
  `officeKitTextRight` scaled guide
- **THEN** PPJ exposes one `customGeometryTextRectangleRightEmu` leaf with
  numeric value `800000`

#### Scenario: Keep a reference-backed right edge opaque

- **WHEN** `a:rect/@r` names a built-in or ordinary declared guide instead of a
  direct numeric token or the exact private literal profile
- **THEN** PPJ does not expose a `customGeometryTextRectangleRightEmu` leaf

### Requirement: Edit only the selected text-rectangle right coordinate

The system SHALL accept a source-bound edit only when the expected right value,
rectangle structure, all four literal edge tokens, private-guide identity (if
present), and requested in-frame ordered coordinate are valid. A direct source
edit MUST replace only `a:rect/@r`; a private-profile edit MUST replace only
the numeric operand of `officeKitTextRight`'s scaled `a:gd/@fmla` while
retaining the `a:rect` reference and scale. The edit MUST preserve left/top/
bottom, guides, paths, handles, connection sites, and package topology. A
stale, out-of-frame, unordered, changed-representation, or unsupported target
MUST fail closed without a partial package.

#### Scenario: Edit and reproject a literal right edge

- **WHEN** a PPJ edit changes a projected right edge from `800000` to `900000`
  with matching source preconditions and the existing rectangle remains
  ordered
- **THEN** only the owning SlidePart changes, the right edge resolves to
  `900000`, and a second projection reports `900000`

#### Scenario: Reject a stale or invalid right-edge edit

- **WHEN** the source literal/guide no longer contains the expected right value
  or the requested coordinate is negative, beyond the shape width, or not
  greater than the left edge
- **THEN** the edit fails closed with diagnostics and does not write the
  source-bound package

#### Scenario: Preserve text-rectangle representation and topology

- **WHEN** the selected right edge is edited
- **THEN** the `officeKitTextRight` representation (when present), left/top/
  bottom edges, guide identity, handles, connection sites, paths, and other
  geometry topology remain unchanged
