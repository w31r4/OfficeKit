## Purpose

This capability preserves and edits one explicit direct rich-text reflection
gradient start position without flattening the surrounding source package.

## ADDED Requirements

### Requirement: Project an imported direct text reflection start position

Projection MUST emit `textReflectionStartPosition` for a direct rich-text run
reflection whose `stPos` is an explicit canonical integer from `0` through
`100000` and whose `endPos` is absent or the canonical value `100000`. The
semantic PPJ value MUST be the position divided by `100000` at
`run.style.reflection.startPosition`.

#### Scenario: Imported run start position becomes a leaf

- **WHEN** a text run `rPr/effectLst/reflection` has `stPos="20000"` and
  `endPos="100000"`
- **THEN** projection exposes a run/text-index-bound leaf whose binding proves
  the raw token `20000` and whose semantic PPJ value is `0.2`

### Requirement: Edit only the direct text reflection start-position token

The source-bound compiler MUST accept a changed canonical position token only
when the same text run and direct reflection owner prove the expected raw
value. It MUST replace only `a:reflection/@stPos` in the owning SlidePart and
MUST preserve all other reflection attributes, effects, text topology, and
package members.

#### Scenario: Run start-position edit preserves the source package

- **WHEN** a projected run start-position leaf changes from `20000` to
  `65000`
- **THEN** compilation changes only the owning SlidePart and re-projection
  returns `startPosition` equal to `0.65`

### Requirement: Keep unsupported direct reflection graphs source-owned

The compiler MUST fail closed for a missing or malformed `stPos`, a
noncanonical or out-of-range token, an `endPos` other than absent or `100000`,
duplicate effect/list owners, unknown attributes or children, and stale leaf
proofs.

#### Scenario: Variable end position is not widened

- **WHEN** the source reflection has `stPos="20000"` and `endPos="80000"`, an
  unknown attribute, or a nonnumeric `stPos`
- **THEN** projection does not issue the start-position leaf and an attempted
  edit is rejected without mutating the source package
