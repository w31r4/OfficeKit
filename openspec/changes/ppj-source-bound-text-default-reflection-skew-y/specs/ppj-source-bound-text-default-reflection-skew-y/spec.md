# Specification: Paragraph default reflection skew Y

## ADDED Requirements

### Requirement: Project an imported paragraph default reflection skew Y

The projection MUST emit `textDefaultReflectionSkewY` for an existing direct
paragraph default-text reflection with a canonical signed `ky` token and MUST
expose its value at `paragraph.style.defaultText.reflection.skewY` using
1/60000-degree precision.

#### Scenario: Imported skew becomes an editable leaf

- **WHEN** a paragraph `pPr/defRPr/effectLst/reflection` has a valid `ky`
  token and full-span positions
- **THEN** projection emits the paragraph-index-bound leaf with the exact raw
  token and the corresponding degree value

### Requirement: Edit only the paragraph default reflection skew token

The source-bound compiler MUST accept a changed canonical `ky` leaf only when
the same paragraph and strict reflection owner prove the expected raw token,
and MUST replace only `a:reflection/@ky` in that SlidePart.

#### Scenario: Skew edit preserves the source package

- **WHEN** a projected paragraph default reflection skew leaf is changed
- **THEN** compilation changes only the owning `SlidePart`, preserves `sx`,
  `sy`, `fadeDir`, `kx`, all other reflection attributes and sibling effects,
  and re-projection returns the new degree value

### Requirement: Keep unsupported reflection transforms source-owned

The compiler MUST fail closed for missing, malformed, stale, out-of-range,
duplicate, child-bearing, or otherwise unsupported paragraph default reflection
graphs, and MUST leave the source package unchanged when an edit is rejected.

#### Scenario: Alignment does not widen the owner

- **WHEN** the reflection contains `algn`, `rotWithShape`, unknown content,
  duplicate effect structure, or a noncanonical `ky`
- **THEN** the strict source-bound proof does not authorize the skew leaf for
  editing (a permissive projection may retain its scalar), and an attempted
  skew edit is rejected or returns the original source without package mutation
