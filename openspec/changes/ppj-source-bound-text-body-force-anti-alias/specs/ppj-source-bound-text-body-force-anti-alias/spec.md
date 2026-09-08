# Text-body force anti-aliasing

## ADDED Requirements

### Requirement: preserve force anti-alias semantics

The presentation wire and PPJ projection MUST preserve a direct
`a:bodyPr/@forceAA` boolean as `textBoxStyle.forceAntiAlias` when the native
text-body profile is otherwise recognized.

#### Scenario: authored force anti-alias round-trip

- **WHEN** authored PPJ supplies `textStyle.forceAntiAlias: true`
- **THEN** the compiler writes `a:bodyPr/@forceAA="1"` and a second
  projection returns `textBoxStyle.forceAntiAlias: true`

### Requirement: issue a bounded native leaf

The source-bound projector MUST issue `textBodyForceAntiAlias` only for one
existing direct `a:bodyPr/@forceAA` token whose value is canonical `0` or `1`
and whose text-body ownership proof succeeds.

#### Scenario: direct token is editable

- **WHEN** a projected text body has one direct canonical `forceAA` token
- **THEN** its native leaf value is a boolean and its native kind is
  `textBodyForceAntiAlias`

#### Scenario: unsupported source stays source-owned

- **WHEN** `forceAA` is absent, duplicated, malformed, inherited, or part of
  an unsupported body-property topology
- **THEN** the projector does not issue `textBodyForceAntiAlias`

### Requirement: splice only the owned attribute

A source-bound edit using `textBodyForceAntiAlias` MUST replace only the
direct `forceAA` token in the owning SlidePart.

#### Scenario: source-bound edit and reprojection

- **WHEN** a caller changes the leaf from `1` to `0`
- **THEN** only the owning SlidePart changes, Open XML remains valid, other
  package parts remain byte-identical, and a second projection returns false
