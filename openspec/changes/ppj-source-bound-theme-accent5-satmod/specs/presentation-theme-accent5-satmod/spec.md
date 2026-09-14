## Purpose

This capability gives PPJ a safe, field-qualified source-bound owner for one
direct accent5 saturation modulation token while preserving the rest of an
imported theme.

## ADDED Requirements

### Requirement: Strict source-bound accent5 satMod projection

The PPTX projector MUST expose
`design.theme.accentTransforms.accent5.satMod` and
`setThemeAccent5SatMod` only when the canonical shared ThemePart has one
direct accent5 RGB leaf with exactly one direct `a:satMod/@val` child, no
extra attributes or descendants, and an integer value in `[0, 100000]`. The
projected value MUST be the native value divided by 100000.

#### Scenario: Project a direct accent5 satMod leaf

- **GIVEN** a canonical imported ThemePart with accent5 `223344` and
  `a:satMod/@val="25000"`
- **WHEN** the presentation is projected to PPJ
- **THEN** the theme contains `accentTransforms.accent5.satMod` equal to
  `0.25`
- **AND** the native reference contains `setThemeAccent5SatMod` authorized
  for `accentTransforms.accent5.satMod`

#### Scenario: Reject unsupported accent5 satMod topology

- **GIVEN** accent5 has a missing, non-RGB, transformed, duplicated, extended,
  or out-of-range satMod topology
- **WHEN** the presentation is projected
- **THEN** the source-bound accent5 satMod field and capability are omitted

### Requirement: Source-bound accent5 satMod editing

The source-preserving compiler MUST accept one authorized change to the
existing `design.theme.accentTransforms.accent5.satMod` field, convert a
finite 0..1 fraction to the DrawingML thousandth integer using away-from-zero
rounding, and change only the existing direct `a:satMod/@val` token in the
canonical ThemePart.

#### Scenario: Edit and reproject one accent5 satMod token

- **GIVEN** a projected direct accent5 satMod owner at `0.25`
- **WHEN** an authorized PPJ program changes it to `0.5`
- **THEN** a no-op compile preserves bytes and reports no changed parts
- **AND** changing it to `0.5` changes only the ThemePart and writes `50000`
- **AND** a fresh projection returns `0.5`

#### Scenario: Fail closed for unsafe accent5 satMod edits

- **GIVEN** a source-bound accent5 satMod owner
- **WHEN** the field is deleted, combined with another theme field,
  accompanied by an unsupported sibling, set outside `[0, 1]`, or its
  capability is altered
- **THEN** compilation fails closed without creating a new native satMod node
