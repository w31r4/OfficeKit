## Purpose

This capability gives PPJ a safe, field-qualified source-bound owner for one
direct accent2 blue-channel modulation token while preserving the rest of an
imported theme.

## ADDED Requirements

### Requirement: Strict source-bound accent2 blueMod projection

The PPTX projector MUST expose
`design.theme.accentTransforms.accent2.blueMod` and
`setThemeAccent2BlueMod` only when the canonical shared ThemePart has one
direct accent2 RGB leaf with exactly one direct `a:blueMod/@val` child, no
extra attributes or descendants, and an integer value in `[0, 100000]`.
The projected value MUST be the native value divided by 100000.

#### Scenario: Project a direct accent2 blueMod leaf

- **GIVEN** a canonical imported ThemePart with accent2 `223344` and
  `a:blueMod/@val="25000"`
- **WHEN** the presentation is projected to PPJ
- **THEN** the theme contains `accentTransforms.accent2.blueMod` equal to
  `0.25`
- **AND** the native reference contains `setThemeAccent2BlueMod` authorized
  for `accentTransforms.accent2.blueMod`

#### Scenario: Reject unsupported accent2 blueMod topology

- **GIVEN** accent2 has a missing, non-RGB, transformed, duplicated, extended,
  or out-of-range blueMod topology
- **WHEN** the presentation is projected
- **THEN** the source-bound accent2 blueMod field and capability are omitted

### Requirement: Source-bound accent2 blueMod editing

The source-preserving compiler MUST accept one authorized change to the
existing `design.theme.accentTransforms.accent2.blueMod` field, convert a
finite 0..1 fraction to the unsigned DrawingML thousandth-percent integer using
away-from-zero rounding, and change only the existing direct `a:blueMod/@val`
token in the canonical ThemePart.

#### Scenario: Edit and reproject one accent2 blueMod token

- **GIVEN** a projected direct accent2 blueMod owner at `0.25`
- **WHEN** an authorized PPJ program changes it to `0.4`
- **THEN** a no-op compile preserves bytes and reports no changed parts
- **AND** changing it to `0.4` changes only the ThemePart and writes `40000`
- **AND** a fresh projection returns `0.4`

#### Scenario: Fail closed for unsafe accent2 blueMod edits

- **GIVEN** a source-bound accent2 blueMod owner
- **WHEN** the field is deleted, combined with another theme field,
  accompanied by an unsupported sibling, set outside `0..1`, or its
  capability is altered
- **THEN** compilation fails closed without creating a new native blueMod node
