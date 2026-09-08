# PPJ authored glow effect

## ADDED Requirements

### Requirement: expose a bounded glow field

The PPJ schema MUST accept a `glow` object on authored shape styles and image
styles. The object MUST contain a `color` and a `radius` in the range `0..1000`
points. Its optional `opacity` MUST be a number in `0..1` or an opacity grammar
token.

#### Scenario: shape glow compiles to the direct effect owner

- **WHEN** an authored shape style declares a blue glow with radius `12` and
  opacity `0.42`
- **THEN** the shape properties contain one `a:effectLst/a:glow` with
  `rad="152400"`, the requested color, and `a:alpha val="42000"`
- **AND** embedded PPJ recovery retains the glow fields

#### Scenario: picture glow uses the same field

- **WHEN** an authored image element declares a glow in its image style
- **THEN** the picture properties contain one direct `a:glow`
- **AND** the glow does not alter the image asset, crop, mask, or border

### Requirement: preserve shadow independence and reject unsafe values

Glow lowering MUST remain independent from the existing outer-shadow field.
When both are present, `a:glow` MUST precede `a:outerShdw` in the effect list.
Non-finite, out-of-range, missing-color, and missing-radius glow values MUST
fail before native output.

#### Scenario: glow and shadow keep separate owners

- **WHEN** an authored shape declares both glow and shadow
- **THEN** the effect list contains exactly one glow and one outer shadow
- **AND** the two color, opacity, and geometry values are not merged

#### Scenario: invalid glow radius fails closed

- **WHEN** a glow declares `radius: 1001`
- **THEN** PPJ validation or native compilation fails

### Requirement: keep broader effects source-owned

The glow field MUST NOT make imported reflection, inner-shadow, soft-edge,
3-D, or arbitrary multi-effect graphs editable. Such graphs MUST remain
source-preserved or fail closed.

#### Scenario: imported broader effects stay opaque

- **WHEN** an imported shape contains a reflection, inner shadow, soft edge,
  3-D effect, or an effect list outside the authored glow-plus-outer-shadow
  profile
- **THEN** projection preserves the source-owned graph or refuses the edit
- **AND** it does not flatten the graph into the PPJ `glow` field
