# Text-body compatible line spacing

## ADDED Requirements

### Requirement: preserve compatible line spacing semantics

The presentation wire and PPJ projection MUST preserve a direct
`a:bodyPr/@compatLnSpc` boolean as
`textBoxStyle.compatibleLineSpacing` when the native text-body profile is
otherwise recognized.

#### Scenario: authored compatible line spacing round-trip

- **WHEN** authored PPJ supplies `textStyle.compatibleLineSpacing: true`
- **THEN** the compiler writes `a:bodyPr/@compatLnSpc="1"` and a second
  projection returns `textBoxStyle.compatibleLineSpacing: true`

### Requirement: issue a bounded native leaf

The source-bound projector MUST issue `textBodyCompatibleLineSpacing` only for
one existing direct `a:bodyPr/@compatLnSpc` token whose value is canonical `0`
or `1` and whose text-body ownership proof succeeds.

#### Scenario: direct token is editable

- **WHEN** a projected text body has one direct canonical `compatLnSpc` token
- **THEN** its native leaf value is a boolean and its native kind is
  `textBodyCompatibleLineSpacing`

#### Scenario: unsupported source stays source-owned

- **WHEN** `compatLnSpc` is absent, duplicated, malformed, inherited, or part
  of an unsupported body-property topology
- **THEN** the projector does not issue `textBodyCompatibleLineSpacing`

### Requirement: splice only the owned attribute

A source-bound edit using `textBodyCompatibleLineSpacing` MUST replace only
the direct `compatLnSpc` token in the owning SlidePart.

#### Scenario: source-bound edit and reprojection

- **WHEN** a caller changes the leaf from `1` to `0`
- **THEN** only the owning SlidePart changes, Open XML remains valid, other
  package parts remain byte-identical, and a second projection returns false
