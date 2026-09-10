## Purpose

This capability preserves and edits one explicit paragraph default reflection
gradient start position without flattening the surrounding source package.

## ADDED Requirements

### Requirement: Project an imported paragraph default reflection start position

Projection MUST emit `textDefaultReflectionStartPosition` for a direct
paragraph default-text reflection whose `stPos` is an explicit canonical
integer from `0` through `100000` and whose `endPos` is absent or the canonical
full-span value `100000`. The semantic PPJ value MUST be the position divided
by `100000` at `paragraph.style.defaultText.reflection.startPosition`.

#### Scenario: Imported start position becomes a leaf

- **WHEN** a paragraph `pPr/defRPr/effectLst/reflection` has `stPos="20000"`
  and no `endPos` or `endPos="100000"`
- **THEN** projection exposes a paragraph-index-bound leaf whose binding proves
  the raw token `20000` and whose semantic PPJ value is `0.2`

### Requirement: Edit only the paragraph default reflection start-position token

The source-bound compiler MUST accept a changed canonical position token only
when the same paragraph and direct reflection owner prove the expected raw
value. It MUST replace only `a:reflection/@stPos` in the owning SlidePart and
MUST preserve all other reflection attributes, effects, topology, and package
members.

#### Scenario: Start position edit preserves the source package

- **WHEN** a projected start-position leaf changes from `20000` to `35000`
- **THEN** compilation changes only the owning SlidePart and re-projection
  returns `startPosition` equal to `0.35`

### Requirement: Keep unsupported reflection graphs source-owned

The compiler MUST fail closed for a missing or malformed `stPos`, a noncanonical
or out-of-range token, an `endPos` other than `100000`, duplicate effect/list
owners, unknown attributes or children, and stale leaf proofs.

#### Scenario: Unsupported position topology is not widened

- **WHEN** the source reflection has `endPos="80000"`, an unknown attribute, or
  a nonnumeric `stPos`
- **THEN** projection does not issue the start-position leaf and an attempted
  edit is rejected without mutating the source package
