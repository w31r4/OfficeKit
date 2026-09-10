# Specification: Paragraph default reflection skew X

## ADDED Requirements

### Requirement: Project an imported paragraph default reflection skew X

The projection MUST emit `textDefaultReflectionSkewX` for an existing direct
paragraph default-text reflection with a canonical signed `kx` token and MUST
expose its value at `paragraph.style.defaultText.reflection.skewX` using
1/60000-degree precision.

#### Scenario: Imported skew becomes an editable leaf

- **WHEN** a paragraph `pPr/defRPr/effectLst/reflection` has a valid `kx`
  token and full-span positions
- **THEN** projection emits the paragraph-index-bound leaf with the exact raw
  token and the corresponding degree value

### Requirement: Edit only the paragraph default reflection skew token

The source-bound compiler MUST accept a changed canonical `kx` leaf only when
the same paragraph and strict reflection owner prove the expected raw token,
and MUST replace only `a:reflection/@kx` in that SlidePart.

#### Scenario: Skew edit preserves the source package

- **WHEN** a projected paragraph default reflection skew leaf is changed
- **THEN** compilation changes only the owning `SlidePart`, preserves `sx`,
  `sy`, `fadeDir`, all other reflection attributes and sibling effects, and
  re-projection returns the new degree value

### Requirement: Keep unsupported reflection transforms source-owned

The compiler MUST fail closed for missing, malformed, stale, out-of-range,
duplicate, child-bearing, or otherwise unsupported paragraph default reflection
graphs, and MUST leave the source package unchanged when an edit is rejected.

#### Scenario: Vertical skew does not widen the owner

- **WHEN** the reflection contains `ky`, `algn`, `rotWithShape`, unknown
  content, duplicate effect structure, or a noncanonical `kx`
- **THEN** projection omits the editable skew leaf and an attempted skew edit is
  rejected or returns the original source without package mutation
