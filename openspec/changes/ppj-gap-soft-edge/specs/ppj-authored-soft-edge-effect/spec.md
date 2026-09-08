# PPJ authored soft-edge effect

## ADDED Requirements

### Requirement: expose a bounded soft-edge field

The PPJ schema MUST accept a `softEdge` object on authored shape styles and
image styles. The object MUST contain a `radius` in the range `0..1000`
points.

#### Scenario: shape soft edge compiles to the direct effect owner

- **WHEN** an authored shape style declares a soft edge with radius `8`
- **THEN** the shape properties contain one `a:softEdge` with `rad="101600"`
- **AND** embedded PPJ recovery retains the soft-edge radius

#### Scenario: picture style soft edge preserves picture owners

- **WHEN** an authored image style declares a soft edge
- **THEN** the picture properties contain one direct `a:softEdge`
- **AND** the soft edge does not alter the image asset, crop, mask, border, or shadow

### Requirement: preserve effect ownership and reject unsupported graphs

Soft-edge lowering MUST remain independent from glow and outer shadow. When
multiple bounded effects are present, the canonical order MUST be `glow`,
`outerShdw`, `softEdge`. Unsupported imported or authored effect children MUST
fail closed rather than being dropped.

#### Scenario: bounded effects keep separate owners

- **WHEN** an authored shape declares glow, outer shadow, and soft edge
- **THEN** the effect list contains one child for each requested effect in the
  canonical order
- **AND** no radius, color, or opacity value is merged between effects

#### Scenario: invalid soft-edge radius fails closed

- **WHEN** a soft edge declares `radius: 1001`
- **THEN** PPJ validation or native compilation fails

#### Scenario: broader imported effects stay opaque

- **WHEN** an imported shape contains a reflection, inner shadow, 3-D effect,
  or an effect list outside the bounded authored profile
- **THEN** projection preserves the source-owned graph or refuses the edit
- **AND** it does not flatten the graph into the PPJ `softEdge` field
