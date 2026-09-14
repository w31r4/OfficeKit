## Purpose

This capability gives PPJ a safe, field-qualified source-bound owner for one
direct accent4 red-channel modulation token while preserving the rest of an
imported theme.

## ADDED Requirements

### Requirement: Strict source-bound accent4 redMod projection

The PPTX projector MUST expose
`design.theme.accentTransforms.accent4.redMod` and
`setThemeAccent4RedMod` only when the canonical shared ThemePart has one
direct accent4 RGB leaf with exactly one direct `a:redMod/@val` child, no
extra attributes or descendants, and an integer value in `[0, 100000]`. The
projected value MUST be the native value divided by 100000.

#### Scenario: Project a direct accent4 redMod leaf

- **GIVEN** a canonical imported ThemePart with accent4 `334455` and
  `a:redMod/@val="25000"`
- **WHEN** the presentation is projected to PPJ
- **THEN** the theme contains `accentTransforms.accent4.redMod` equal to
  `0.25`
- **AND** the native reference contains `setThemeAccent4RedMod` authorized
  for `accentTransforms.accent4.redMod`

#### Scenario: Reject unsupported accent4 redMod topology

- **GIVEN** accent4 has a missing, non-RGB, transformed, duplicated, extended,
  or out-of-range redMod topology
- **WHEN** the presentation is projected
- **THEN** the source-bound accent4 redMod field and capability are omitted

### Requirement: Source-bound accent4 redMod editing

The source-preserving compiler MUST accept one authorized change to the
existing `design.theme.accentTransforms.accent4.redMod` field, convert a
finite 0..1 fraction to the DrawingML thousandth integer using away-from-zero
rounding, and change only the existing direct `a:redMod/@val` token in the
canonical ThemePart.

#### Scenario: Edit and reproject one accent4 redMod token

- **GIVEN** a projected direct accent4 redMod owner at `0.25`
- **WHEN** an authorized PPJ program changes it to `0.5`
- **THEN** a no-op compile preserves bytes and reports no changed parts
- **AND** changing it to `0.5` changes only the ThemePart and writes `50000`
- **AND** a fresh projection returns `0.5`

#### Scenario: Fail closed for unsafe accent4 redMod edits

- **GIVEN** a source-bound accent4 redMod owner
- **WHEN** the field is deleted, combined with another theme field,
  accompanied by an unsupported sibling, set outside `[0, 1]`, or its
  capability is altered
- **THEN** compilation fails closed without creating a new native redMod node
