# Text-warp preset

## ADDED Requirements

### Requirement: preserve the bounded text-warp preset

The presentation wire and PPJ projection MUST preserve one direct
`a:bodyPr/a:prstTxWarp/@prst` value as `textBoxStyle.textWarpPreset` when the
preset is in the DrawingML `TextShapeValues` vocabulary and the element is
bare or has one recognized literal adjustment list.

#### Scenario: authored text-warp preset round-trip

- **WHEN** authored PPJ supplies `textStyle.textWarpPreset: "textArchUp"`
- **THEN** the compiler writes one `a:prstTxWarp/@prst="textArchUp"` and a
  second projection returns the same PPJ field.

### Requirement: issue a bounded native leaf

The source-bound projector MUST issue `textBodyWarpPreset` only for one
existing direct `a:prstTxWarp` child with exactly one canonical `prst`
attribute, a supported preset value, and either no child or one recognized
literal `a:avLst/a:gd` adjustment list. Unsupported adjustment or extension
topology suppresses the bounded preset leaf.

#### Scenario: direct preset is editable

- **WHEN** a projected text body has one direct `a:prstTxWarp/@prst` preset
- **THEN** its native leaf value is the canonical preset string and its native
  kind is `textBodyWarpPreset`

#### Scenario: unsupported transform stays source-owned

- **WHEN** the preset is absent, duplicated, malformed, outside the finite
  vocabulary, or has an unsupported adjustment/extension child
- **THEN** the projector does not issue `textBodyWarpPreset`

### Requirement: splice only the owned preset attribute

A source-bound edit using `textBodyWarpPreset` MUST replace only the direct
`prst` token in the owning SlidePart.

#### Scenario: source-bound edit and reprojection

- **WHEN** a caller changes the leaf from `textArchUp` to `textPlain`
- **THEN** only the owning SlidePart changes, text and other package parts
  remain byte-identical, Open XML remains valid, and a second projection
  returns `textPlain`.
