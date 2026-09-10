## Purpose

This capability preserves and edits one explicit paragraph default reflection
gradient end position without flattening the surrounding source package.

## ADDED Requirements

### Requirement: Project an imported paragraph default reflection end position

Projection MUST emit `textDefaultReflectionEndPosition` for a direct paragraph
default-text reflection whose `endPos` is an explicit canonical integer from
`0` through `100000` and whose `startPos` is absent or the canonical value
`0`. The semantic PPJ value MUST be the position divided by `100000` at
`paragraph.style.defaultText.reflection.endPosition`.

#### Scenario: Imported end position becomes a leaf

- **WHEN** a paragraph `pPr/defRPr/effectLst/reflection` has
  `endPos="80000"` and no `startPos` or `startPos="0"`
- **THEN** projection exposes a paragraph-index-bound leaf whose binding proves
  the raw token `80000` and whose semantic PPJ value is `0.8`

### Requirement: Edit only the paragraph default reflection end-position token

The source-bound compiler MUST accept a changed canonical position token only
when the same paragraph and direct reflection owner prove the expected raw
value. It MUST replace only `a:reflection/@endPos` in the owning SlidePart and
MUST preserve all other reflection attributes, effects, topology, and package
members.

#### Scenario: End position edit preserves the source package

- **WHEN** a projected end-position leaf changes from `80000` to `65000`
- **THEN** compilation changes only the owning SlidePart and re-projection
  returns `endPosition` equal to `0.65`

### Requirement: Keep unsupported reflection graphs source-owned

The compiler MUST fail closed for a missing or malformed `endPos`, a
noncanonical or out-of-range token, a `startPos` other than `0`, duplicate
effect/list owners, unknown attributes or children, and stale leaf proofs.

#### Scenario: Two variable endpoints are not widened

- **WHEN** the source reflection has `stPos="20000"` and `endPos="80000"`,
  an unknown attribute, or a nonnumeric `endPos`
- **THEN** projection does not issue the end-position leaf and an attempted
  edit is rejected without mutating the source package
