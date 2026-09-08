## ADDED Requirements

### Requirement: expose a bounded authored text inner shadow

The PPJ schema MUST allow `textStyle.innerShadow` on a rich-text run style and
on a paragraph `defaultText` style. It MUST use the existing inner-shadow
profile: a direct RGB or bounded theme color, `blur`, `distance`, and `angle`,
plus optional numeric or opacity-token `opacity`.

#### Scenario: authored run inner shadow compiles to the direct owner

- **WHEN** an authored rich-text run declares an inner shadow with direct RGB
  color `#16324F`, blur `5`, distance `2`, angle `45`, and opacity `0.36`
- **THEN** the run properties contain one `a:innerShdw` with the requested
  color, geometry, and `a:alpha val="36000"`
- **AND** embedded PPJ recovery retains the inner-shadow fields

#### Scenario: paragraph default inner shadow uses `a:defRPr`

- **WHEN** a paragraph `style.defaultText` declares a bounded inner shadow
- **THEN** the compiler writes it to the paragraph's `a:defRPr`
- **AND** the effect does not become duplicated direct formatting on every run

### Requirement: preserve text effect ownership and reject unsupported graphs

Inner-shadow lowering MUST remain independent from text glow and outer shadow.
When multiple bounded effects are present, the canonical order MUST be
`glow`, `innerShdw`, `outerShdw`. Unsupported effect children MUST fail closed
rather than being dropped.

#### Scenario: bounded text effects keep separate owners

- **WHEN** an authored run or default run declares glow, inner shadow, and
  outer shadow
- **THEN** its effect list contains one child for each requested effect in the
  canonical order
- **AND** no color, opacity, or geometry value is merged between effects

#### Scenario: invalid inner-shadow geometry fails closed

- **WHEN** a text inner shadow declares a negative blur/distance, an invalid
  angle, or opacity outside the bounded range
- **THEN** PPJ validation or native compilation fails before output promotion

#### Scenario: broader imported effects stay opaque

- **WHEN** an imported text run contains reflection, 3-D, or an arbitrary
  multi-effect graph
- **THEN** projection preserves the source-owned graph or refuses the edit
- **AND** it does not flatten the graph into PPJ `innerShadow`
