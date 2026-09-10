# Specification: Paragraph default reflection scale Y

## ADDED Requirements

### Requirement: Project an imported paragraph default reflection scale Y

The projection MUST emit `textDefaultReflectionScaleY` for an existing direct
paragraph default-text reflection with a canonical signed `sy` token, and MUST
expose its value at `paragraph.style.defaultText.reflection.scaleY` using the
native 1/100000 ratio unit.

#### Scenario: Imported scale becomes an editable leaf

- **WHEN** a paragraph `pPr/defRPr/effectLst/reflection` has a valid `sy`
  token and full-span positions
- **THEN** projection emits the paragraph-index-bound leaf with the exact raw
  token and the corresponding PPJ ratio

### Requirement: Edit only the paragraph default reflection scale token

The source-bound compiler MUST accept a changed canonical `sy` leaf only when
the same paragraph and strict reflection owner prove the expected raw token,
and MUST replace only `a:reflection/@sy` in that SlidePart.

#### Scenario: Scale edit preserves the source package

- **WHEN** a projected paragraph default reflection scale leaf is changed
- **THEN** compilation changes only the owning SlidePart, preserves `sx`,
  `fadeDir`, all other reflection attributes and sibling effects, and
  re-projection returns the new ratio

### Requirement: Keep unsupported reflection scale sources source-owned

The compiler MUST fail closed for missing, malformed, stale, out-of-range,
duplicate, child-bearing, or otherwise unsupported paragraph default reflection
graphs, and MUST leave the source package unchanged when an edit is rejected.

#### Scenario: Other reflection transforms do not widen the owner

- **WHEN** the reflection contains `kx`, `ky`, `algn`, `rotWithShape`, unknown
  content, duplicate effect structure, or a noncanonical `sy`
- **THEN** projection omits the editable scale leaf and an attempted scale edit
  is rejected or returns the original source without package mutation
