# PPJ formal text language owner

## ADDED Requirements

### Requirement: include language in formal text precedence

The authored PPJ compiler MUST accept `text.language` as a bounded
`stylePrecedence` target. When the rule is present, it MUST select the first
declared source containing `language` and lower that BCP-47 value to the
direct DrawingML run language owner.

#### Scenario: theme language reaches an authored run

- **WHEN** a run omits `language`, the theme text style declares a BCP-47
  language, and `stylePrecedence` lists `theme` before `default`
- **THEN** the authored run contains the theme language in `a:rPr/@lang`
- **AND** a second projection recovers the same language value

### Requirement: preserve historical lookup without a rule

When no `text.language` precedence rule is declared, the compiler MUST keep
the existing direct/default text-style lookup behavior and MUST NOT require a
new source or wire field.

#### Scenario: inline language remains authoritative

- **WHEN** a run directly declares `language` and no formal language rule is
  present
- **THEN** the direct run language is emitted and recovered unchanged

### Requirement: keep source-bound inheritance closed

The source-bound compiler MUST reject a newly declared formal language owner
over an imported source, while an existing direct `fontLanguage` native leaf
remains governed by its current revision-bound capability.

#### Scenario: source-bound formal owner is rejected

- **WHEN** a source-bound PPJ adds a `text.language` precedence rule
- **THEN** validation fails with the existing formal text-owner boundary
- **AND** no inherited language owner is guessed or rewritten
