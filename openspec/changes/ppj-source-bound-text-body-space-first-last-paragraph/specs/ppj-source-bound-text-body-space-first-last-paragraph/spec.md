# Text-body first/last paragraph spacing

## ADDED Requirements

### Requirement: preserve first/last paragraph spacing semantics

The presentation wire and PPJ projection MUST preserve a direct
`a:bodyPr/@spcFirstLastPara` boolean as
`textBoxStyle.spaceFirstLastParagraph` when the native text-body profile is
otherwise recognized.

#### Scenario: authored first/last paragraph spacing round-trip

- **WHEN** authored PPJ supplies `textStyle.spaceFirstLastParagraph: true`
- **THEN** the compiler writes `a:bodyPr/@spcFirstLastPara="1"` and a second
  projection returns `textBoxStyle.spaceFirstLastParagraph: true`

### Requirement: issue a bounded native leaf

The source-bound projector MUST issue `textBodySpaceFirstLastParagraph` only
for one existing direct `a:bodyPr/@spcFirstLastPara` token whose value is
canonical `0` or `1` and whose text-body ownership proof succeeds.

#### Scenario: direct token is editable

- **WHEN** a projected text body has one direct canonical
  `spcFirstLastPara` token
- **THEN** its native leaf value is a boolean and its native kind is
  `textBodySpaceFirstLastParagraph`

#### Scenario: unsupported source stays source-owned

- **WHEN** `spcFirstLastPara` is absent, duplicated, malformed, inherited, or
  part of an unsupported body-property topology
- **THEN** the projector does not issue
  `textBodySpaceFirstLastParagraph`

### Requirement: splice only the owned attribute

A source-bound edit using `textBodySpaceFirstLastParagraph` MUST replace only
the direct `spcFirstLastPara` token in the owning SlidePart.

#### Scenario: source-bound edit and reprojection

- **WHEN** a caller changes the leaf from `1` to `0`
- **THEN** only the owning SlidePart changes, Open XML remains valid, other
  package parts remain byte-identical, and a second projection returns false
