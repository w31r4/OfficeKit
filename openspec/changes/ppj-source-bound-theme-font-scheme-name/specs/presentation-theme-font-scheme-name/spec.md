## ADDED Requirements

### Requirement: Project an existing font scheme name

The projector MUST expose design.theme.fontScheme.name only when the canonical shared ThemePart contains one non-empty a:fontScheme/@name attribute, and MUST issue setThemeFontSchemeName for exactly fontScheme.name.

#### Scenario: Existing name is projected

- GIVEN an imported presentation has one canonical ThemePart with a non-empty a:fontScheme/@name
- WHEN the presentation is projected to PPJ
- THEN design.theme.fontScheme.name equals the source value
- AND the theme native reference contains setThemeFontSchemeName with fontScheme.name
- AND the other font slots and theme graph remain source-owned

#### Scenario: Missing name stays opaque

- GIVEN the canonical font scheme has no non-empty name attribute
- WHEN the presentation is projected
- THEN fontScheme.name and its capability are omitted
- AND the projector does not create a missing OOXML attribute

### Requirement: Preserve a source-bound name edit

The compiler MUST accept one non-empty replacement for design.theme.fontScheme.name only when the projected capability is present, and MUST patch only the existing a:fontScheme/@name attribute.

#### Scenario: One name edit changes only the ThemePart

- GIVEN a projected source-bound PPJ contains fontScheme.name
- WHEN the name changes and all other PPJ fields remain unchanged
- THEN compilation succeeds
- AND the changed-part set contains only the canonical ThemePart
- AND a fresh projection returns the replacement name

#### Scenario: Unchanged name is byte stable

- GIVEN a projected source-bound PPJ is compiled without edits
- WHEN the package is exported
- THEN the source package bytes remain unchanged

#### Scenario: Invalid name edits fail closed

- GIVEN a projected source-bound PPJ
- WHEN the name is deleted or empty, changed together with another font scheme slot, or submitted with a forged capability
- THEN compilation fails
- AND no package is emitted
