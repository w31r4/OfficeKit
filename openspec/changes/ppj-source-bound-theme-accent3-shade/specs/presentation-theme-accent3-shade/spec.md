## Purpose

This capability gives PPJ a safe, field-qualified source-bound owner for one
direct accent3 shade token while preserving the rest of an imported theme.

## ADDED Requirements

### Requirement: Strict source-bound accent3 shade projection

The PPTX projector MUST expose
`design.theme.accentTransforms.accent3.shade` and
`setThemeAccent3Shade` only when the canonical shared ThemePart has one direct
accent3 RGB leaf with exactly one direct `a:shade/@val` child, no extra
attributes or descendants, and an integer value in `[0, 100000]`. The
projected value MUST be the native value divided by 100000.

#### Scenario: Project a direct accent3 shade leaf

- **GIVEN** a canonical imported ThemePart with accent3 `334455` and
  `a:shade/@val="35000"`
- **WHEN** the presentation is projected to PPJ
- **THEN** the theme contains `accentTransforms.accent3.shade` equal to `0.35`
- **AND** the native reference contains `setThemeAccent3Shade` authorized for
  `accentTransforms.accent3.shade`

#### Scenario: Reject unsupported accent3 shade topology

- **GIVEN** accent3 has a missing, non-RGB, transformed, duplicated, extended,
  or out-of-range shade topology
- **WHEN** the presentation is projected
- **THEN** the source-bound accent3 shade field and capability are omitted

### Requirement: Source-bound accent3 shade editing

The source-preserving compiler MUST accept one authorized change to the
existing `design.theme.accentTransforms.accent3.shade` field, convert a finite
0..1 fraction to the DrawingML 1/100000 integer using away-from-zero
rounding, and change only the existing direct `a:shade/@val` token in the
canonical ThemePart.

#### Scenario: Edit and reproject one accent3 shade token

- **GIVEN** a projected direct accent3 shade owner at `0.35`
- **WHEN** an authorized PPJ program changes it to `0.7`
- **THEN** a no-op compile preserves bytes and reports no changed parts
- **AND** changing it to `0.7` changes only the ThemePart and writes `70000`
- **AND** a fresh projection returns `0.7`

#### Scenario: Fail closed for unsafe accent3 shade edits

- **GIVEN** a source-bound accent3 shade owner
- **WHEN** the field is deleted, combined with another theme field,
  accompanied by an unsupported sibling, set outside `0..1`, or its capability
  is altered
- **THEN** compilation fails closed without creating a new native shade node
