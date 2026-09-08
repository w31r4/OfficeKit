# Text-warp adjustments

## ADDED Requirements

### Requirement: preserve literal text-warp adjustments

The presentation wire and PPJ projection MUST preserve the ordered
`a:prstTxWarp/a:avLst/a:gd` adjustment values as
`textBoxStyle.textWarpAdjustments` when the direct warp child has one
supported preset and every guide is an attribute-only `name` plus canonical
`fmla="val N"` entry with a unique name.

#### Scenario: authored text-warp adjustment round-trip

- **WHEN** authored PPJ supplies `textStyle.textWarpPreset: "textArchUp"` and
  `textStyle.textWarpAdjustments: [{"name":"adj","value":50000}]`
- **THEN** the compiler writes one direct `a:prstTxWarp` with `prst="textArchUp"`
  and one `a:gd name="adj" fmla="val 50000"`, and a second projection returns
  both fields.

### Requirement: issue indexed adjustment leaves

The source-bound projector MUST issue `textBodyWarpAdjustment` for each
literal adjustment in the one recognized direct warp list. The native leaf
index MUST identify the ordered guide position, and the leaf value MUST be the
signed integer value rather than the formula text.

#### Scenario: literal adjustment is editable

- **WHEN** a projected text body contains one canonical `a:gd` adjustment
- **THEN** it exposes the matching `textBodyWarpAdjustment` leaf with the
  integer value and preserves `textWarpPreset` plus the ordered adjustment
  entry.

#### Scenario: unsupported adjustment remains source-owned

- **WHEN** the list is absent, empty, formula-driven, extension-bearing,
  duplicate-named, malformed, or contains unsupported child topology
- **THEN** the projector does not issue `textBodyWarpAdjustment` for that
  warp list.

### Requirement: splice one adjustment value

A source-bound edit using `textBodyWarpAdjustment` MUST replace only the
selected direct `a:gd/@fmla` value token in the owning SlidePart.

#### Scenario: source-bound adjustment edit and reprojection

- **WHEN** a caller changes the indexed leaf from `50000` to `60000`
- **THEN** only the owning SlidePart changes, the preset/name/order and other
  package parts remain byte-identical, Open XML remains valid, and a second
  projection returns `60000`.
