# Calculated custom-geometry adjustment formula

## ADDED Requirements

### Requirement: expose one bounded calculated adjustment formula

The system MUST expose a `customGeometryAdjustmentFormula` native leaf only for
a direct ordered `a:avLst/a:gd` adjustment with a non-empty name, no children
or extension attributes, and a canonical non-`val N` formula accepted by the
existing bounded custom-geometry formula graph.

#### Scenario: project a calculated adjustment without flattening the graph

- **WHEN** an imported custom geometry contains a recognized adjustment list
  with a calculated adjustment formula
- **THEN** the projected shape exposes the formula string at that adjustment
  index while guides, handles, connection sites, paths, and the rest of the
  graph remain source-owned

#### Scenario: edit only the selected adjustment formula token

- **WHEN** a source-bound PPJ edit changes an issued formula leaf to another
  canonical formula accepted by the same graph
- **THEN** only that adjustment's `fmla` value changes in the owning SlidePart,
  the other geometry XML remains unchanged, and a subsequent projection reports
  the new formula

#### Scenario: reject unsupported graph changes

- **WHEN** the formula is literal, noncanonical, malformed, unsupported,
  extension-bearing, child-bearing, partial-image-fill, or would invalidate the
  surrounding graph
- **THEN** the system MUST keep the formula source-owned or fail closed rather
  than issue or apply a partial graph edit
