# Text-body WordArt marker

## ADDED Requirements

### Requirement: preserve the direct WordArt marker

The presentation wire and PPJ projection MUST preserve a direct
`a:bodyPr/@fromWordArt` boolean as `textBoxStyle.fromWordArt` when the native
text-body profile is otherwise recognized.

#### Scenario: authored WordArt marker round-trip

- **WHEN** authored PPJ supplies `textStyle.fromWordArt: true`
- **THEN** the compiler writes `a:bodyPr/@fromWordArt="1"` and a second
  projection returns `textBoxStyle.fromWordArt: true`

### Requirement: issue a bounded native leaf

The source-bound projector MUST issue `textBodyFromWordArt` only for one
existing direct `a:bodyPr/@fromWordArt` token whose value is canonical `0` or
`1` and whose text-body ownership proof succeeds.

#### Scenario: direct token is editable

- **WHEN** a projected text body has one direct canonical `fromWordArt` token
- **THEN** its native leaf value is a boolean and its native kind is
  `textBodyFromWordArt`

#### Scenario: unsupported source stays source-owned

- **WHEN** `fromWordArt` is absent, duplicated, malformed, inherited, or part
  of an unsupported WordArt/body-property topology
- **THEN** the projector does not issue `textBodyFromWordArt`

### Requirement: splice only the owned attribute

A source-bound edit using `textBodyFromWordArt` MUST replace only the direct
`fromWordArt` token in the owning SlidePart.

#### Scenario: source-bound edit and reprojection

- **WHEN** a caller changes the leaf from `1` to `0`
- **THEN** only the owning SlidePart changes, Open XML remains valid, other
  package parts remain byte-identical, and a second projection returns false
