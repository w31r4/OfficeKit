# Specification: Paragraph default reflection alignment

## ADDED Requirements

### Requirement: Project an imported paragraph default reflection alignment

The projection MUST emit `textDefaultReflectionAlignment` for an existing direct
paragraph default-text reflection with a canonical `algn` token and MUST expose
its value at `paragraph.style.defaultText.reflection.alignment`. The token MUST
be one of `tl`, `t`, `tr`, `l`, `ctr`, `r`, `bl`, `b`, or `br`.

#### Scenario: Imported alignment becomes an editable leaf

- **WHEN** a paragraph `pPr/defRPr/effectLst/reflection` has a valid `algn`
  token and full-span positions
- **THEN** projection emits the paragraph-index-bound leaf with the exact token
  and the corresponding PPJ alignment value

### Requirement: Edit only the paragraph default reflection alignment token

The source-bound compiler MUST accept a changed canonical `algn` leaf only when
the same paragraph and strict reflection owner prove the expected raw token, and
MUST replace only `a:reflection/@algn` in that SlidePart.

#### Scenario: Alignment edit preserves the source package

- **WHEN** a projected paragraph default reflection alignment leaf is changed
- **THEN** compilation changes only the owning `SlidePart`, preserves all other
  reflection attributes, sibling effects, paragraph/run topology, and ZIP
  members, and re-projection returns the new alignment value

### Requirement: Keep unsupported reflection graphs source-owned

The compiler MUST fail closed for missing, malformed, stale, out-of-range,
duplicate, child-bearing, rotate-with-shape, or otherwise unsupported paragraph
default reflection graphs, and MUST leave the source package unchanged when an
edit is rejected.

#### Scenario: Unsupported content does not widen the owner

- **WHEN** the reflection contains `rotWithShape`, unknown content, duplicate
  effect structure, or a noncanonical `algn`
- **THEN** the strict source-bound proof does not authorize the alignment leaf
  for editing, and an attempted edit is rejected or returns the original source
  without package mutation
