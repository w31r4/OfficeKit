## Purpose

This capability preserves and edits one explicit direct rich-text reflection
fade direction without flattening the surrounding source package.

## ADDED Requirements

### Requirement: Project an imported direct text reflection fade angle

Projection MUST emit `textReflectionFadeAngleDegrees` for a direct rich-text run
whose reflection has an explicit canonical `fadeDir` integer from `0` through
`21599999`, no scale/skew/alignment/rotate-with-shape transform, and absent or
full-span endpoint positions. The semantic PPJ value MUST be the token divided
by `60000` at `run.style.reflection.fadeAngle`.

#### Scenario: Imported run fade direction becomes a leaf

- **WHEN** a text run `rPr/effectLst/reflection` has `fadeDir="2730000"`
  and full-span positions
- **THEN** projection exposes a run/text-index-bound leaf whose binding proves
  the raw token and whose semantic PPJ value is `45.5`

### Requirement: Edit only the direct text reflection fade token

The source-bound compiler MUST accept a changed canonical fade token only when
the same text run and direct reflection owner prove the expected raw value. It
MUST replace only `a:reflection/@fadeDir` in the owning SlidePart and MUST
preserve all other reflection attributes, effects, text topology, and package
members.

#### Scenario: Run fade-angle edit preserves the source package

- **WHEN** a projected fade-angle leaf changes from `2730000` to `5415000`
- **THEN** compilation changes only the owning SlidePart and re-projection
  returns `fadeAngle` equal to `90.25`

### Requirement: Keep unsupported direct reflection transforms source-owned

The compiler MUST fail closed for a missing or malformed `fadeDir`, a
noncanonical or out-of-range token, variable endpoints, scale/skew/alignment/
rotate-with-shape transforms, duplicate effect/list owners, unknown attributes
or children, and stale leaf proofs.

#### Scenario: Unsupported transform is not widened

- **WHEN** the source reflection also has `kx`, an unknown attribute, or a
  nonnumeric `fadeDir`
- **THEN** projection does not issue the fade-angle leaf and an attempted edit
  is rejected without mutating the source package
