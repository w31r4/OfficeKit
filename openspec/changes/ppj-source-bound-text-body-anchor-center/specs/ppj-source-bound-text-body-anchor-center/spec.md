# Text-body anchor centering

## ADDED Requirements

### Requirement: preserve anchor-center semantics

The presentation wire and PPJ projection MUST preserve a direct
`a:bodyPr/@anchorCtr` boolean as `textBoxStyle.anchorCenter` when the native
text-body profile is otherwise recognized.

#### Scenario: authored anchor-center round-trip

- **WHEN** authored PPJ supplies `textStyle.anchorCenter: true`
- **THEN** the compiler writes `a:bodyPr/@anchorCtr="1"` and a second
  projection returns `textBoxStyle.anchorCenter: true`

### Requirement: issue a bounded native leaf

The source-bound projector MUST issue `textBodyAnchorCenter` only for one
existing direct `a:bodyPr/@anchorCtr` token whose value is canonical `0` or
`1` and whose text-body ownership proof succeeds.

#### Scenario: direct token is editable

- **WHEN** a projected text body has one direct canonical `anchorCtr` token
- **THEN** its native leaf value is a boolean and its native kind is
  `textBodyAnchorCenter`

#### Scenario: unsupported source stays source-owned

- **WHEN** `anchorCtr` is absent, duplicated, malformed, inherited, or part of
  an unsupported body-property topology
- **THEN** the projector does not issue `textBodyAnchorCenter`

### Requirement: splice only the owned attribute

A source-bound edit using `textBodyAnchorCenter` MUST replace only the direct
`anchorCtr` token in the owning SlidePart.

#### Scenario: source-bound edit and reprojection

- **WHEN** a caller changes the leaf from `1` to `0`
- **THEN** only the owning SlidePart changes, Open XML remains valid, other
  package parts remain byte-identical, and a second projection returns false
