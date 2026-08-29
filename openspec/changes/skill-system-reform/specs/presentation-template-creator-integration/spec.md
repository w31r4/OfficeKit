## ADDED Requirements

### Requirement: Creator calibration shall use the shared primitive language
The Presentation Template Creator MUST load the shared primitive and typography
references when creating calibration pages, while keeping its published output
to style guidance, original visual evidence, and the existing packaging schema.

#### Scenario: Creator builds an image-led calibration page
- **WHEN** the reference style requires native background, scene ordering, or
  chart/image composition
- **THEN** the Creator may use the corresponding OfficeKit primitive and record
  the observed capability, but does not publish the source deck or code

### Requirement: Creator output shall not promise runtime completeness
Published template guidance MUST distinguish reusable visual evidence from
third-party source-graph editability and MUST not claim that examples are fixed
layouts or a complete OOXML profile.

#### Scenario: User selects a style template
- **WHEN** Presentations loads the template Skill
- **THEN** it derives a deck-specific grammar and composes freely, with source
  continuation handled by the separate imported-edit route
