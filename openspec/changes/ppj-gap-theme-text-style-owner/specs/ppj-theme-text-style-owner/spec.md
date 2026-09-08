## Purpose

This capability gives the formal PPJ text precedence profile an explicit,
authored theme fallback owner that can be resolved and reviewed without
claiming arbitrary PowerPoint theme XML editing.

## ADDED Requirements

### Requirement: Theme text-style owner

The PPJ schema MAY declare `design.theme.textStyle` as a direct authored
text-style owner. When a formal `stylePrecedence` rule for
`text.size`, `text.bold`, `text.italic`, `text.font`, or `text.fontFamily`
declares `theme`, the compiler and reviewer MUST consult that owner and MUST
respect the first declared source that has a value.

#### Scenario: Theme supplies the first available value

- **GIVEN** a formal text scalar rule declares `theme` before `default`
- **AND** `design.theme.textStyle` supplies that scalar
- **WHEN** the authored program is compiled or reviewed
- **THEN** the effective scalar MUST come from `theme`
- **AND** the review resolution MUST identify `theme` as the winning source

#### Scenario: Theme falls through to default

- **GIVEN** a formal text scalar rule declares `theme` before `default`
- **AND** `design.theme.textStyle` omits that scalar
- **AND** the grammar declares a matching default token
- **WHEN** the authored program is compiled or reviewed
- **THEN** the default token MUST supply the effective scalar

### Requirement: Authored-only theme lowering

The authored compiler MUST lower a selected theme scalar into the existing
native text-run owner. It MUST NOT create or claim a source-bound editable
PowerPoint theme style graph from this field.

#### Scenario: Effective theme value survives projection

- **GIVEN** a valid authored PPJ whose theme supplies a formal text scalar
- **WHEN** it is compiled, the embedded PPJ is removed, and the PPTX is
  projected back to PPJ
- **THEN** the effective native run scalar MUST be recoverable
- **AND** the projection MUST NOT advertise the discarded theme declaration as
  source-bound editable

#### Scenario: Source-bound theme declaration fails closed

- **GIVEN** a source-bound PPJ contains `design.theme.textStyle`
- **WHEN** it is submitted for compilation against its source package
- **THEN** compilation MUST fail closed with a diagnostic identifying the
  unsupported source-bound theme owner
