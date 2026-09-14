## Purpose

This capability gives PPJ a safe, field-qualified source-bound owner for one
direct accent3 hue-offset token while preserving the rest of an imported theme.

## ADDED Requirements

### Requirement: Strict source-bound accent3 hueOff projection

The PPTX projector MUST expose
`design.theme.accentTransforms.accent3.hueOff` and
`setThemeAccent3HueOff` only when the canonical shared ThemePart has one
direct accent3 RGB leaf with exactly one direct `a:hueOff/@val` child, no
extra attributes or descendants, and an integer value in
`[-21600000, 21600000]`. The projected value MUST be the native value divided
by 60000.

#### Scenario: Project a direct accent3 hueOff leaf

- **GIVEN** a canonical imported ThemePart with accent3 `334455` and
  `a:hueOff/@val="12600000"`
- **WHEN** the presentation is projected to PPJ
- **THEN** the theme contains `accentTransforms.accent3.hueOff` equal to
  `210`
- **AND** the native reference contains `setThemeAccent3HueOff` authorized
  for `accentTransforms.accent3.hueOff`

#### Scenario: Reject unsupported accent3 hueOff topology

- **GIVEN** accent3 has a missing, non-RGB, transformed, duplicated, extended,
  or out-of-range hueOff topology
- **WHEN** the presentation is projected
- **THEN** the source-bound accent3 hueOff field and capability are omitted

### Requirement: Source-bound accent3 hueOff editing

The source-preserving compiler MUST accept one authorized change to the
existing `design.theme.accentTransforms.accent3.hueOff` field, convert a
finite -360..360 degree offset to the DrawingML 1/60000-degree integer using
away-from-zero rounding, and change only the existing direct `a:hueOff/@val`
token in the canonical ThemePart.

#### Scenario: Edit and reproject one accent3 hueOff token

- **GIVEN** a projected direct accent3 hueOff owner at `210`
- **WHEN** an authorized PPJ program changes it to `-120`
- **THEN** a no-op compile preserves bytes and reports no changed parts
- **AND** changing it to `-120` changes only the ThemePart and writes
  `-7200000`
- **AND** a fresh projection returns `-120`

#### Scenario: Fail closed for unsafe accent3 hueOff edits

- **GIVEN** a source-bound accent3 hueOff owner
- **WHEN** the field is deleted, combined with another theme field,
  accompanied by an unsupported sibling, set outside `-360..360`, or its
  capability is altered
- **THEN** compilation fails closed without creating a new native hueOff node
