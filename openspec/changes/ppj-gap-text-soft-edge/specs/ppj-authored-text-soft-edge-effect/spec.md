## ADDED Requirements

### Requirement: expose a bounded authored text soft edge

The PPJ schema MUST allow `textStyle.softEdge` on a rich-text run style and on
a paragraph `defaultText` style. It MUST contain a `radius` in the existing
`0..1000` point range.

#### Scenario: authored run soft edge compiles to the direct owner

- **WHEN** an authored rich-text run declares soft-edge radius `8`
- **THEN** the run properties contain one `a:softEdge` with radius `50800`
- **AND** embedded PPJ recovery retains the soft-edge radius

#### Scenario: paragraph default soft edge uses `a:defRPr`

- **WHEN** a paragraph `style.defaultText` declares a bounded soft edge
- **THEN** the compiler writes it to the paragraph's `a:defRPr`
- **AND** the effect does not become duplicated direct formatting on every run

### Requirement: preserve text effect ownership and reject unsupported graphs

Soft-edge lowering MUST remain independent from the other bounded text
effects. When multiple bounded effects are present, the canonical order MUST
be `glow`, `innerShdw`, `outerShdw`, `reflection`, `softEdge`. Unsupported
effect children MUST fail closed rather than being dropped.

#### Scenario: bounded text effects keep separate owners

- **WHEN** an authored run or default run declares glow, inner shadow, outer
  shadow, reflection, and soft edge
- **THEN** its effect list contains one child for each requested effect in the
  canonical order
- **AND** no radius, color, opacity, geometry, or position value is merged
  between effects

#### Scenario: invalid soft-edge radius fails closed

- **WHEN** a text soft edge declares a radius outside `0..1000` points
- **THEN** PPJ validation or native compilation fails before output promotion

#### Scenario: broader imported effects stay opaque

- **WHEN** an imported text run contains 3-D or an arbitrary multi-effect graph
- **THEN** projection preserves the source-owned graph or refuses the edit
- **AND** it does not flatten the graph into PPJ `softEdge`
