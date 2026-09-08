# PPJ authored reflection effect

## ADDED Requirements

### Requirement: expose a bounded reflection field

The PPJ schema MUST accept a `reflection` object on authored shape styles and
image styles. It MUST contain `blur`, `startOpacity`, `endOpacity`, `distance`,
and `angle`. Opacity values MAY be numeric or opacity-token references.

#### Scenario: shape reflection compiles to the direct effect owner

- **WHEN** an authored shape style declares blur `4`, start opacity `0.45`,
  end opacity `0.08`, distance `3`, and angle `90`
- **THEN** the shape properties contain one `a:reflection` with the requested
  geometry and alpha values
- **AND** the reflection uses the deterministic full-span positions
- **AND** embedded PPJ recovery retains the reflection fields

#### Scenario: picture style reflection preserves picture owners

- **WHEN** an authored image style declares a reflection
- **THEN** the picture properties contain one direct `a:reflection`
- **AND** the effect does not alter the image asset, crop, mask, border, or
  other bounded effect fields

### Requirement: preserve effect ownership and reject unsupported graphs

Reflection lowering MUST remain independent from glow, inner shadow, outer
shadow, and soft edge. When multiple bounded effects are present, the canonical
order MUST be `glow`, `innerShdw`, `outerShdw`, `reflection`, `softEdge`.
Unsupported effect children MUST fail closed rather than being dropped.

#### Scenario: bounded effects keep separate owners

- **WHEN** an authored shape declares glow, inner shadow, outer shadow,
  reflection, and soft edge
- **THEN** the effect list contains one child for each requested effect in the
  canonical order
- **AND** no opacity or geometry value is merged between effects

#### Scenario: invalid reflection geometry fails closed

- **WHEN** a reflection declares `blur: -1`, an opacity outside `0..1`, or an
  angle outside the existing shadow range
- **THEN** PPJ validation or native compilation fails

#### Scenario: broader imported effects stay opaque

- **WHEN** an imported shape contains reflection variants, 3-D, or an
  arbitrary multi-effect graph
- **THEN** projection preserves the source-owned graph or refuses the edit
- **AND** it does not flatten the graph into PPJ `reflection`
