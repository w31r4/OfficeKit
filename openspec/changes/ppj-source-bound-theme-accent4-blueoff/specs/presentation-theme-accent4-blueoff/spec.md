## Purpose

This capability gives PPJ a safe, field-qualified source-bound owner for one
direct accent4 blue-channel offset token while preserving the rest of an
imported theme.

## ADDED Requirements

### Requirement: Strict source-bound accent4 blueOff projection

The PPTX projector MUST expose
`design.theme.accentTransforms.accent4.blueOff` and
`setThemeAccent4BlueOff` only when the canonical shared ThemePart has one
direct accent4 RGB leaf with exactly one direct `a:blueOff/@val` child, no
extra attributes or descendants, and an integer value in `[-100000, 100000]`.
The projected value MUST be the native value divided by 100000.

#### Scenario: Project a direct accent4 blueOff leaf

- **GIVEN** a canonical imported ThemePart with accent4 `334455` and
  `a:blueOff/@val="-25000"`
- **WHEN** the presentation is projected to PPJ
- **THEN** the theme contains `accentTransforms.accent4.blueOff` equal to
  `-0.25`
- **AND** the native reference contains `setThemeAccent4BlueOff` authorized
  for `accentTransforms.accent4.blueOff`

#### Scenario: Reject unsupported accent4 blueOff topology

- **GIVEN** accent4 has a missing, non-RGB, transformed, duplicated, extended,
  or out-of-range blueOff topology
- **WHEN** the presentation is projected
- **THEN** the source-bound accent4 blueOff field and capability are omitted

### Requirement: Source-bound accent4 blueOff editing

The source-preserving compiler MUST accept one authorized change to the
existing `design.theme.accentTransforms.accent4.blueOff` field, convert a
finite -1..1 fraction to the DrawingML signed thousandth integer using
away-from-zero rounding, and change only the existing direct `a:blueOff/@val`
token in the canonical ThemePart.

#### Scenario: Edit and reproject one accent4 blueOff token

- **GIVEN** a projected direct accent4 blueOff owner at `-0.25`
- **WHEN** an authorized PPJ program changes it to `0.4`
- **THEN** a no-op compile preserves bytes and reports no changed parts
- **AND** changing it to `0.4` changes only the ThemePart and writes `40000`
- **AND** a fresh projection returns `0.4`

#### Scenario: Fail closed for unsafe accent4 blueOff edits

- **GIVEN** a source-bound accent4 blueOff owner
- **WHEN** the field is deleted, combined with another theme field,
  accompanied by an unsupported sibling, set outside `-1..1`, or its
  capability is altered
- **THEN** compilation fails closed without creating a new native blueOff node
