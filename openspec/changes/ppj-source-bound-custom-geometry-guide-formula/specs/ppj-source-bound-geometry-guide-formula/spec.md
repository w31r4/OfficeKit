# Calculated custom-geometry guide formula

## ADDED Requirements

### Requirement: expose one bounded calculated guide formula

The system MUST expose a `customGeometryGuideFormula` native leaf only for a
direct ordered `a:gdLst/a:gd` guide with a unique non-empty name, no children or
extension attributes, and a canonical non-`val N` formula accepted by the
existing bounded custom-geometry formula graph.

#### Scenario: project a calculated guide without flattening the graph

- **WHEN** an imported custom geometry contains a recognized guide list with a
  calculated guide formula
- **THEN** the projected shape exposes the formula string at that guide index
  while handles, connection sites, paths, and the rest of the guide graph
  remain source-owned

#### Scenario: edit only the selected formula token

- **WHEN** a source-bound PPJ edit changes an issued formula leaf to another
  canonical formula accepted by the same graph
- **THEN** only that guide's `fmla` value changes in the owning SlidePart, the
  other geometry XML remains unchanged, and a subsequent projection reports
  the new formula

#### Scenario: reject unsupported graph changes

- **WHEN** the formula is literal, noncanonical, malformed, unsupported,
  duplicate-named, extension-bearing, child-bearing, or would invalidate the
  surrounding graph
- **THEN** the system MUST keep the formula source-owned or fail closed rather
  than issue or apply a partial graph edit
