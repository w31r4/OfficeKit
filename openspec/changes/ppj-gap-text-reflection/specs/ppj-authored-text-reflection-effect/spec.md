## ADDED Requirements

### Requirement: expose a bounded authored text reflection

The PPJ schema MUST allow `textStyle.reflection` on a rich-text run style and
on a paragraph `defaultText` style. It MUST use the existing reflection
profile: bounded `blur`, `startOpacity`, `endOpacity`, `distance`, and `angle`.

#### Scenario: authored run reflection compiles to the direct owner

- **WHEN** an authored rich-text run declares blur `5`, start opacity `0.42`,
  end opacity `0.08`, distance `12`, and angle `45`
- **THEN** the run properties contain one `a:reflection` with the requested
  geometry, `stPos="0"`, and `endPos="100000"`
- **AND** embedded PPJ recovery retains the reflection fields

#### Scenario: paragraph default reflection uses `a:defRPr`

- **WHEN** a paragraph `style.defaultText` declares a bounded reflection
- **THEN** the compiler writes it to the paragraph's `a:defRPr`
- **AND** the effect does not become duplicated direct formatting on every run

### Requirement: preserve text effect ownership and reject unsupported graphs

Reflection lowering MUST remain independent from text glow, inner shadow, and
outer shadow. When multiple bounded text effects are present, the canonical
order MUST be `glow`, `innerShdw`, `outerShdw`, `reflection`. Unsupported effect
children MUST fail closed rather than being dropped.

#### Scenario: bounded text effects keep separate owners

- **WHEN** an authored run or default run declares glow, inner shadow, outer
  shadow, and reflection
- **THEN** its effect list contains one child for each requested effect in the
  canonical order
- **AND** no color, opacity, geometry, or position value is merged between
  effects

#### Scenario: invalid reflection geometry fails closed

- **WHEN** a text reflection declares invalid opacity, negative distance, or an
  angle outside the bounded range
- **THEN** PPJ validation or native compilation fails before output promotion

#### Scenario: broader imported effects stay opaque

- **WHEN** an imported text run contains an unsupported reflection variant, 3-D,
  or an arbitrary multi-effect graph
- **THEN** projection preserves the source-owned graph or refuses the edit
- **AND** it does not flatten the graph into PPJ `reflection`
