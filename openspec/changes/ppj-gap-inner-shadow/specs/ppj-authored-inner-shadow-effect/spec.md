# PPJ authored inner-shadow effect

## ADDED Requirements

### Requirement: expose a bounded inner-shadow field

The PPJ schema MUST accept an `innerShadow` object on authored shape styles
and image styles. It MUST contain `color`, `blur`, `distance`, and `angle` with
the existing shadow bounds, and MAY contain numeric or opacity-token `opacity`.

#### Scenario: shape inner shadow compiles to the direct effect owner

- **WHEN** an authored shape style declares an inner shadow with a direct RGB
  color, blur `5`, distance `2`, angle `45`, and opacity `0.36`
- **THEN** the shape properties contain one `a:innerShdw` with the requested
  color, geometry, and `a:alpha val="36000"`
- **AND** embedded PPJ recovery retains the inner-shadow fields

#### Scenario: picture style inner shadow preserves picture owners

- **WHEN** an authored image style declares an inner shadow
- **THEN** the picture properties contain one direct `a:innerShdw`
- **AND** the effect does not alter the image asset, crop, mask, border, or
  soft-edge field

### Requirement: preserve effect ownership and reject unsupported graphs

Inner-shadow lowering MUST remain independent from glow, outer shadow, and
soft edge. When multiple bounded effects are present, the canonical order MUST
be `glow`, `innerShdw`, `outerShdw`, `softEdge`. Unsupported effect children
MUST fail closed rather than being dropped.

#### Scenario: bounded effects keep separate owners

- **WHEN** an authored shape declares glow, inner shadow, outer shadow, and
  soft edge
- **THEN** the effect list contains one child for each requested effect in the
  canonical order
- **AND** no color, opacity, or geometry value is merged between effects

#### Scenario: invalid inner-shadow geometry fails closed

- **WHEN** an inner shadow declares `blur: -1` or an angle outside the existing
  shadow range
- **THEN** PPJ validation or native compilation fails

#### Scenario: broader imported effects stay opaque

- **WHEN** an imported shape contains reflection, 3-D, or an arbitrary
  multi-effect graph
- **THEN** projection preserves the source-owned graph or refuses the edit
- **AND** it does not flatten the graph into PPJ `innerShadow`
