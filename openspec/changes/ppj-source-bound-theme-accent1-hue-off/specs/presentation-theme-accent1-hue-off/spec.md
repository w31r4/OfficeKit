# Presentation theme accent1 hueOff

## ADDED Requirements

### Requirement: Strict source-bound accent1 hueOff projection

The PPTX projector MUST expose `design.theme.accentTransforms.accent1.hueOff`
and `setThemeAccent1HueOff` only when the canonical shared ThemePart has one
direct accent1 RGB leaf with exactly one direct `a:hueOff/@val` child, no extra
attributes or descendants, and an integer value in
`[-21600000, 21600000]`. The projected value MUST be the native value divided
by 60000.

#### Scenario: Project a direct hueOff leaf

- **GIVEN** a canonical imported ThemePart with accent1 `112233` and
  `a:hueOff/@val="750000"`
- **WHEN** the presentation is projected to PPJ
- **THEN** the theme contains `accentTransforms.accent1.hueOff` equal to
  `12.5`
- **AND** the native reference contains `setThemeAccent1HueOff` authorized for
  `accentTransforms.accent1.hueOff`

#### Scenario: Reject unsupported hueOff topology

- **GIVEN** accent1 has a missing, non-RGB, transformed, duplicated, extended,
  or out-of-range hueOff topology
- **WHEN** the presentation is projected
- **THEN** the source-bound hueOff field and capability are omitted

### Requirement: Source-bound accent1 hueOff editing

The source-preserving compiler MUST accept one authorized change to the
existing `accentTransforms.accent1.hueOff` field, convert degrees to the
DrawingML 1/60000-degree integer using away-from-zero rounding, and change only
the existing direct `a:hueOff/@val` token in the canonical ThemePart. The value
MUST be finite and within `[-360, 360]` degrees.

#### Scenario: Edit and reproject one hueOff token

- **GIVEN** a projected direct hueOff owner at `12.5` degrees
- **WHEN** an authorized PPJ program changes it to `12.5` degrees
- **THEN** a no-op compile preserves bytes and reports no changed parts
- **AND** changing it to `30` degrees changes only the ThemePart and writes
  `1800000`
- **AND** a fresh projection returns `30` degrees

#### Scenario: Fail closed for unsafe edits

- **GIVEN** a source-bound hueOff owner
- **WHEN** the field is deleted, combined with another theme field, accompanied
  by an unsupported sibling, set outside `[-360, 360]`, or its capability is
  altered
- **THEN** compilation fails closed without creating a new native hueOff node
