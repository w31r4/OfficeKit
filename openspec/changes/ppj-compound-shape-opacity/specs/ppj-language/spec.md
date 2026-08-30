## ADDED Requirements

### Requirement: Authored PPJ shapes support compound visual opacity

The authored PPJ compiler SHALL treat `shape.style.opacity` as a multiplier on
every directly owned visible branch of that shape. It SHALL preserve branch-
local alpha, native editability and the single PPJ element identity.

#### Scenario: Compound translucent shape

- **WHEN** a source-free PPJ shape declares opacity below one together with a
  gradient or image fill, outline, shadow and explicitly painted text
- **THEN** each branch receives the multiplied effective alpha
- **AND** the output remains one native editable shape
- **AND** the embedded PPJ recovers the original semantic opacity exactly

#### Scenario: Unrepresentable text branch

- **WHEN** a translucent shape contains visible text whose paint remains
  inherited or contains a highlight without a bounded alpha representation
- **THEN** compilation fails before writing output with a path-specific error
