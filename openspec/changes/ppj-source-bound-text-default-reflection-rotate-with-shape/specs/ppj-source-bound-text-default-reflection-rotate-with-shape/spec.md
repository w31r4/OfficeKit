# Specification: paragraph default reflection rotate with shape

## ADDED Requirements

### Requirement: Project an imported paragraph default reflection rotate flag

The projection MUST emit `textDefaultReflectionRotateWithShape` for an
existing direct paragraph default-text reflection with an explicit canonical
`rotWithShape` token and MUST expose its boolean value at
`paragraph.style.defaultText.reflection.rotateWithShape`.

#### Scenario: Imported rotate flag becomes an editable leaf

- **WHEN** a paragraph `pPr/defRPr/effectLst/reflection` has full-span positions
  and `rotWithShape="1"` or `rotWithShape="0"`
- **THEN** projection emits the paragraph-index-bound boolean leaf with the
  corresponding PPJ value

### Requirement: Edit only the paragraph default reflection rotate token

The source-bound compiler MUST accept a changed canonical boolean leaf only when
the same paragraph and strict reflection owner prove the expected raw token, and
MUST replace only `a:reflection/@rotWithShape` in that SlidePart.

#### Scenario: Rotate flag edit preserves the source package

- **WHEN** a projected paragraph default reflection rotate leaf changes from
  true to false or false to true
- **THEN** compilation changes only the owning `SlidePart`, preserves all other
  reflection attributes, sibling effects, paragraph/run topology, and ZIP
  members, and re-projection returns the new boolean

### Requirement: Keep unsupported reflection graphs source-owned

The compiler MUST fail closed for missing, malformed, stale, duplicate,
child-bearing, non-full-span, unknown, or otherwise unsupported paragraph
reflection graphs, and MUST leave the source package unchanged when an edit is
rejected.

#### Scenario: Unknown transform does not widen the owner

- **WHEN** the reflection contains an unknown transform attribute or unknown
  child instead of the supported canonical profile
- **THEN** the strict source-bound proof does not authorize the rotate leaf, and
  an attempted edit is rejected or returns the original source without package
  mutation
