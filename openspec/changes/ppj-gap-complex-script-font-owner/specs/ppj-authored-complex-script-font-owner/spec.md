# PPJ authored complex-script font owner

## ADDED Requirements

### Requirement: expose a direct complex-script typeface

The PPJ text style grammar MUST accept `fontFamilyComplexScript` wherever the
bounded direct text style is accepted. A literal or string grammar token MUST
resolve to a non-empty typeface of at most 255 characters, and the native
compiler MUST lower it to `a:cs/@typeface` without changing omitted font
owners.

#### Scenario: authored run keeps three font owners

- **WHEN** a PPJ text run declares Latin, East Asian, and complex-script font
  families
- **THEN** the authored `a:rPr` contains one matching `a:latin`, `a:ea`, and
  `a:cs` child
- **AND** a second projection recovers all three fields

### Requirement: project and edit a simple imported owner

The native projection MUST issue `fontFamilyComplexScript` for a simple direct
run `a:cs` child. A source-bound edit of that leaf MUST preserve surrounding
run content and unrelated native children while changing only its typeface.

#### Scenario: source-bound single-leaf edit

- **WHEN** an imported run has one direct typeface-only `a:cs` child and a
  revision-bound edit changes its complex-script family
- **THEN** the output changes only the proven `a:cs/@typeface`
- **AND** re-import exposes the requested complex-script family

### Requirement: preserve the closed boundary

The codec MUST reject empty or overlong typefaces and MUST leave child-bearing,
ambiguous, inherited, or effect-bearing complex-script font graphs outside
the bounded editing profile.

#### Scenario: unsupported native graph stays closed

- **WHEN** an imported `a:cs` has extra attributes, children, or duplicate
  direct owners
- **THEN** the codec does not issue an editable complex-script leaf
