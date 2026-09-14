## Purpose

This capability gives PPJ a safe, field-qualified source-bound owner for one
direct accent3 tint token while preserving the rest of an imported theme.

## ADDED Requirements

### Requirement: Strict source-bound accent3 tint projection

The PPTX projector MUST expose
`design.theme.accentTransforms.accent3.tint` and
`setThemeAccent3Tint` only when the canonical shared ThemePart has one direct
accent3 RGB leaf with exactly one direct `a:tint/@val` child, no extra
attributes or descendants, and an integer value in `[0, 100000]`. The
projected value MUST be the native value divided by 100000.

#### Scenario: Project a direct accent3 tint leaf

- **GIVEN** a canonical imported ThemePart with accent3 `334455` and
  `a:tint/@val="35000"`
- **WHEN** the presentation is projected to PPJ
- **THEN** the theme contains `accentTransforms.accent3.tint` equal to `0.35`
- **AND** the native reference contains `setThemeAccent3Tint` authorized for
  `accentTransforms.accent3.tint`

#### Scenario: Reject unsupported accent3 tint topology

- **GIVEN** accent3 has a missing, non-RGB, transformed, duplicated, extended,
  or out-of-range tint topology
- **WHEN** the presentation is projected
- **THEN** the source-bound accent3 tint field and capability are omitted

### Requirement: Source-bound accent3 tint editing

The source-preserving compiler MUST accept one authorized change to the
existing `design.theme.accentTransforms.accent3.tint` field, convert a finite
0..1 fraction to the DrawingML 1/100000 integer using away-from-zero
rounding, and change only the existing direct `a:tint/@val` token in the
canonical ThemePart.

#### Scenario: Edit and reproject one accent3 tint token

- **GIVEN** a projected direct accent3 tint owner at `0.35`
- **WHEN** an authorized PPJ program changes it to `0.7`
- **THEN** a no-op compile preserves bytes and reports no changed parts
- **AND** changing it to `0.7` changes only the ThemePart and writes `70000`
- **AND** a fresh projection returns `0.7`

#### Scenario: Fail closed for unsafe accent3 tint edits

- **GIVEN** a source-bound accent3 tint owner
- **WHEN** the field is deleted, combined with another theme field,
  accompanied by an unsupported sibling, set outside `0..1`, or its capability
  is altered
- **THEN** compilation fails closed without creating a new native tint node
