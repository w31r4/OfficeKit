## ADDED Requirements

### Requirement: Project an imported paragraph default reflection fade angle

The projection MUST emit a `textDefaultReflectionFadeAngleDegrees` native leaf
for an existing direct paragraph default-text reflection whose `fadeDir` token
is valid, and MUST expose that value at
`paragraph.style.defaultText.reflection.fadeAngle` in degrees using the native
1/60000-degree unit.

#### Scenario: Imported fade direction becomes an editable leaf

- **WHEN** a paragraph `pPr/defRPr/effectLst/reflection` has one canonical
  `fadeDir` token in the supported range
- **THEN** projection emits the paragraph-index-bound leaf with the exact
  integer token and the PPJ reflection contains the corresponding degree value

### Requirement: Edit only the paragraph default reflection fade token

The source-bound compiler MUST accept a changed valid fade-angle leaf only when
the same paragraph, effect list, and reflection still prove the expected raw
token, and MUST replace only `a:reflection/@fadeDir` in that SlidePart.

#### Scenario: Fade angle edit preserves the source package

- **WHEN** a projected paragraph default reflection fade leaf is changed
- **THEN** compilation changes only the owning slide part, replaces only
  `fadeDir`, preserves sibling effects and other reflection attributes, and
  re-projection returns the new degree value

### Requirement: Keep unsupported reflection sources source-owned

The compiler MUST fail closed for a missing, malformed, out-of-range, stale,
duplicate, or otherwise unsupported paragraph default reflection graph, and
MUST leave the source package byte-for-byte unchanged when the edit is rejected.

#### Scenario: Unsupported fade source cannot be widened implicitly

- **WHEN** the reflection has no valid `fadeDir`, has an invalid/bound token,
  has duplicate effect structure, or has unsupported child/attribute content
- **THEN** projection omits the editable fade leaf and an attempted fade edit
  is rejected or returned as the original source without a package mutation
