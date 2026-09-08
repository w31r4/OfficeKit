## Purpose

Gives the existing East Asian text typeface field the same bounded, observable
formal precedence ownership as the other supported text run scalars.

## ADDED Requirements

### Requirement: East Asian font precedence is explicit

An authored PPJ formal precedence profile MAY declare
`text.fontFamilyEastAsia`. When declared, the value MUST resolve from the
listed supported owners in order: run, paragraph, element, style reference,
layout, master, theme, or default. A resolved value MUST be emitted as the
East Asian typeface of the authored text run. If no listed owner provides a
value, the existing Latin font fallback MAY continue to supply the East Asian
typeface.

#### Scenario: Theme supplies the East Asian font

- **WHEN** a formal rule includes `theme` and `design.theme.textStyle` provides
  `fontFamilyEastAsia`
- **THEN** authored compilation and read-only review report the theme value as
  the winning East Asian typeface

#### Scenario: Default supplies the East Asian font

- **WHEN** a formal rule includes `default`, no higher owner provides an East
  Asian typeface, and the grammar has a default East Asian font token
- **THEN** the default token value is emitted and reported as the winner

### Requirement: East Asian precedence stays bounded and source-safe

The rule MUST use the existing supported text owner vocabulary and string/font
value validation. Source-bound PPJ MUST NOT gain a capability to introduce or
rewrite formal text precedence owners or a native theme font scheme through
this field.

#### Scenario: Unsupported source is rejected

- **WHEN** a formal East Asian font rule names an unsupported source
- **THEN** PPJ validation fails with the existing formal precedence diagnostic

#### Scenario: Source-bound formal owner is rejected

- **WHEN** a source-bound PPJ carries a formal owner declaration or new theme
  text-style owner for the East Asian font
- **THEN** the request fails closed rather than editing an inferred native
  inheritance graph

### Requirement: Existing text behavior remains compatible

When no `text.fontFamilyEastAsia` rule is declared, authored text MUST retain
the existing direct/inline/default behavior, including fallback from a resolved
Latin `fontFamily`. Existing formal rules for the other text scalars MUST keep
their prior source order and output.

#### Scenario: Existing profile omits the target

- **WHEN** a program has formal size/bold rules but no East Asian font rule
- **THEN** its East Asian typeface follows the pre-existing direct field or
  Latin-font fallback path

#### Scenario: Existing formal scalar remains unchanged

- **WHEN** a program declares both `text.size` and `text.fontFamilyEastAsia`
- **THEN** size still follows its own declared source list independently of the
  East Asian font resolution
