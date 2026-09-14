## Purpose

This capability gives PPJ a safe, field-qualified source-bound owner for one
 direct accent3 hue-modulation token while preserving the rest of an imported theme.

## ADDED Requirements

### Requirement: Strict source-bound accent3 hueMod projection

The PPTX projector MUST expose
`design.theme.accentTransforms.accent3.hueMod` and
`setThemeAccent3HueMod` only when the canonical shared ThemePart has one
direct accent3 RGB leaf with exactly one direct `a:hueMod/@val` child, no
extra attributes or descendants, and an integer value in `[0, 100000]`.
The projected value MUST be the native value divided by 100000.

#### Scenario: Project a direct accent3 hueMod leaf

- **GIVEN** a canonical imported ThemePart with accent3 `334455` and
  `a:hueMod/@val="35000"`
- **WHEN** the presentation is projected to PPJ
- **THEN** the theme contains `accentTransforms.accent3.hueMod` equal to
  `0.35`
- **AND** the native reference contains `setThemeAccent3HueMod` authorized
  for `accentTransforms.accent3.hueMod`

#### Scenario: Reject unsupported accent3 hueMod topology

- **GIVEN** accent3 has a missing, non-RGB, transformed, duplicated, extended,
  or out-of-range hueMod topology
- **WHEN** the presentation is projected
- **THEN** the source-bound accent3 hueMod field and capability are omitted

### Requirement: Source-bound accent3 hueMod editing

The source-preserving compiler MUST accept one authorized change to the
existing `design.theme.accentTransforms.accent3.hueMod` field, convert a
finite 0..1 fraction to the DrawingML thousandth-percent integer using
away-from-zero rounding, and change only the existing direct `a:hueMod/@val`
token in the canonical ThemePart.

#### Scenario: Edit and reproject one accent3 hueMod token

- **GIVEN** a projected direct accent3 hueMod owner at `0.35`
- **WHEN** an authorized PPJ program changes it to `0.7`
- **THEN** a no-op compile preserves bytes and reports no changed parts
- **AND** changing it to `0.7` changes only the ThemePart and writes `70000`
- **AND** a fresh projection returns `0.7`

#### Scenario: Fail closed for unsafe accent3 hueMod edits

- **GIVEN** a source-bound accent3 hueMod owner
- **WHEN** the field is deleted, combined with another theme field,
  accompanied by an unsupported sibling, set outside `0..1`, or its
  capability is altered
- **THEN** compilation fails closed without creating a new native hueMod node
