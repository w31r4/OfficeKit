## Purpose

This capability gives PPJ a bounded source-bound way to edit one literal left
edge of a custom-geometry text rectangle while retaining its representation,
identity, and surrounding topology.

## ADDED Requirements

### Requirement: Project a literal text-rectangle left edge

The system SHALL expose one `customGeometryTextRectangleLeftEmu` native leaf
when a recognized custom geometry has exactly one text rectangle whose four
resolved edges are canonical non-negative integers within the shape-local
frame, the rectangle is ordered, and its left edge is literal. A direct
numeric `a:rect/@l` token or the exact OfficeKit literal profile
`l="officeKitTextLeft"` with a canonical scaled `a:gd` formula MAY provide the
left token. Reference-backed, missing, malformed, extension-bearing,
negative, out-of-frame, or unordered rectangles MUST remain source-owned.

#### Scenario: Project an authored private-profile left edge

- **WHEN** an authored/imported custom geometry has a valid text rectangle with
  a left edge of `300000` shape-local EMUs represented by the OfficeKit
  `officeKitTextLeft` scaled guide
- **THEN** PPJ exposes one `customGeometryTextRectangleLeftEmu` leaf with
  numeric value `300000`

#### Scenario: Keep a reference-backed left edge opaque

- **WHEN** `a:rect/@l` names a built-in or ordinary declared guide instead of a
  direct numeric token or the exact private literal profile
- **THEN** PPJ does not expose a `customGeometryTextRectangleLeftEmu` leaf

### Requirement: Edit only the selected text-rectangle left coordinate

The system SHALL accept a source-bound edit only when the expected left value,
rectangle structure, all four literal edge tokens, private-guide identity (if
present), and requested in-frame ordered coordinate are valid. A direct source
edit MUST replace only `a:rect/@l`; a private-profile edit MUST replace only
the numeric operand of `officeKitTextLeft`'s scaled `a:gd/@fmla` while retaining
the `a:rect` reference and scale. The edit MUST preserve top/right/bottom,
guides, paths, handles, connection sites, and package topology. A stale,
out-of-frame, unordered, changed-representation, or unsupported target MUST
fail closed without a partial package.

#### Scenario: Edit and reproject a literal left edge

- **WHEN** a PPJ edit changes a projected left edge from `300000` to `420000`
  with matching source preconditions and the existing rectangle remains
  ordered
- **THEN** only the owning SlidePart changes, the left edge resolves to
  `420000`, and a second projection reports `420000`

#### Scenario: Reject a stale or invalid left-edge edit

- **WHEN** the source literal/guide no longer contains the expected left value
  or the requested coordinate is negative, beyond the shape width, or not less
  than the right edge
- **THEN** the edit fails closed with diagnostics and does not write the
  source-bound package

#### Scenario: Preserve text-rectangle representation and topology

- **WHEN** the selected left edge is edited
- **THEN** the `officeKitTextLeft` representation (when present), top/right/
  bottom edges, guide identity, handles, connection sites, paths, and other
  geometry topology remain unchanged
