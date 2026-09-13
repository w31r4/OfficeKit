## Purpose

This capability gives PPJ a safe, field-qualified source-bound owner for one
direct accent2 alpha offset token while preserving the rest of an imported
theme.

## ADDED Requirements

### Requirement: Strict source-bound accent2 alphaOff projection

The PPTX projector MUST expose
`design.theme.accentTransforms.accent2.alphaOff` and
`setThemeAccent2AlphaOff` only when the canonical shared ThemePart has one
direct accent2 RGB leaf with exactly one direct `a:alphaOff/@val` child, no
extra attributes or descendants, and an integer value in `[-100000, 100000]`.
The projected value MUST be the native value divided by 100000.

#### Scenario: Project a direct accent2 alphaOff leaf

- **GIVEN** a canonical imported ThemePart with accent2 `223344` and
  `a:alphaOff/@val="-25000"`
- **WHEN** the presentation is projected to PPJ
- **THEN** the theme contains `accentTransforms.accent2.alphaOff` equal to
  `-0.25`
- **AND** the native reference contains `setThemeAccent2AlphaOff` authorized
  for `accentTransforms.accent2.alphaOff`

#### Scenario: Reject unsupported accent2 alphaOff topology

- **GIVEN** accent2 has a missing, non-RGB, transformed, duplicated, extended,
  or out-of-range alphaOff topology
- **WHEN** the presentation is projected
- **THEN** the source-bound accent2 alphaOff field and capability are omitted

### Requirement: Source-bound accent2 alphaOff editing

The source-preserving compiler MUST accept one authorized change to the
existing `design.theme.accentTransforms.accent2.alphaOff` field, convert a
finite -1..1 fraction to the DrawingML thousandth integer using away-from-zero
rounding, and change only the existing direct `a:alphaOff/@val` token in the
canonical ThemePart.

#### Scenario: Edit and reproject one accent2 alphaOff token

- **GIVEN** a projected direct accent2 alphaOff owner at `-0.25`
- **WHEN** an authorized PPJ program changes it to `0.5`
- **THEN** a no-op compile preserves bytes and reports no changed parts
- **AND** changing it to `0.5` changes only the ThemePart and writes `50000`
- **AND** a fresh projection returns `0.5`

#### Scenario: Fail closed for unsafe accent2 alphaOff edits

- **GIVEN** a source-bound accent2 alphaOff owner
- **WHEN** the field is deleted, combined with another theme field,
  accompanied by an unsupported sibling, set outside `[-1, 1]`, or its
  capability is altered
- **THEN** compilation fails closed without creating a new native alphaOff node
