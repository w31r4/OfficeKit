## Purpose

This capability preserves and edits one explicit direct rich-text reflection
gradient end position without flattening the surrounding source package.

## ADDED Requirements

### Requirement: Project an imported direct text reflection end position

Projection MUST emit `textReflectionEndPosition` for a direct rich-text run
reflection whose `endPos` is an explicit canonical integer from `0` through
`100000` and whose `stPos` is absent or the canonical value `0`. The semantic
PPJ value MUST be the position divided by `100000` at
`run.style.reflection.endPosition`.

#### Scenario: Imported run end position becomes a leaf

- **WHEN** a text run `rPr/effectLst/reflection` has `stPos="0"` and
  `endPos="80000"`
- **THEN** projection exposes a run/text-index-bound leaf whose binding proves
  the raw token `80000` and whose semantic PPJ value is `0.8`

### Requirement: Edit only the direct text reflection end-position token

The source-bound compiler MUST accept a changed canonical position token only
when the same text run and direct reflection owner prove the expected raw
value. It MUST replace only `a:reflection/@endPos` in the owning SlidePart and
MUST preserve all other reflection attributes, effects, text topology, and
package members.

#### Scenario: Run end-position edit preserves the source package

- **WHEN** a projected run end-position leaf changes from `80000` to `65000`
- **THEN** compilation changes only the owning SlidePart and re-projection
  returns `endPosition` equal to `0.65`

### Requirement: Keep unsupported direct reflection graphs source-owned

The compiler MUST fail closed for a missing or malformed `endPos`, a
noncanonical or out-of-range token, a `stPos` other than absent or `0`,
duplicate effect/list owners, unknown attributes or children, and stale leaf
proofs.

#### Scenario: Variable start position is not widened

- **WHEN** the source reflection has `stPos="20000"` and `endPos="80000"`, an
  unknown attribute, or a nonnumeric `endPos`
- **THEN** projection does not issue the end-position leaf and an attempted
  edit is rejected without mutating the source package
