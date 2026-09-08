# Flat-text Z coordinate

## ADDED Requirements

### Requirement: preserve a bounded flat-text coordinate

The presentation wire and PPJ projection MUST preserve a direct
`a:flatTx/@z` coordinate as `textBoxStyle.flatTextZ` when the child has one
canonical signed integer `z` in the bounded signed 32-bit range.

#### Scenario: authored flat-text coordinate round-trip

- **WHEN** authored PPJ supplies `textStyle.flatTextZ: 1200`
- **THEN** the compiler writes one direct `a:flatTx z="1200"` and a second
  projection returns `flatTextZ: 1200`.

### Requirement: issue the source-bound coordinate leaf

The source-bound projector MUST issue `textBodyFlatTextZ` for one recognized
direct `a:flatTx/@z`, with its numeric value and no arbitrary XML selector.

#### Scenario: canonical coordinate is editable

- **WHEN** an imported text body contains one direct canonical `flatTx z`
- **THEN** it exposes one matching `textBodyFlatTextZ` leaf and preserves
  surrounding text and 3D markup.

#### Scenario: unsupported coordinate remains source-owned

- **WHEN** `flatTx` is missing `z`, duplicated, extra-attribute-bearing,
  child-bearing, noncanonical, or outside the bounded range
- **THEN** the projector does not issue `textBodyFlatTextZ`.

### Requirement: splice only the coordinate token

A source-bound `textBodyFlatTextZ` edit MUST replace only the selected direct
`a:flatTx/@z` value token in the owning SlidePart.

#### Scenario: source-bound edit and reprojection

- **WHEN** a caller changes the coordinate from `1200` to `2400`
- **THEN** only the owning SlidePart changes, all sibling markup and other ZIP
  parts remain byte-identical, Open XML remains valid, and a second projection
  returns `2400`.
