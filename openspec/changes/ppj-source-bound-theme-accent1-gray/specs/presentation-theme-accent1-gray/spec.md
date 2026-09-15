## Purpose

Expose one existing direct accent1 grayscale transform as a field-qualified
source-bound PPJ owner while preserving the imported theme graph.

## ADDED Requirements

### Requirement: Strict source-bound accent1 gray projection

The PPTX projector MUST expose
`design.theme.accentTransforms.accent1.gray` and
`setThemeAccent1Gray` only when the canonical shared ThemePart has one direct
accent1 RGB leaf with exactly one direct `a:gray` child, no extra attributes or
descendants, and no alpha or sibling transform. The projected field MUST be
`true`.

#### Scenario: Project a direct accent1 gray leaf

- **GIVEN** an imported presentation has one shared ThemePart whose accent1
  color is `112233` followed by one direct `a:gray` child
- **WHEN** the presentation is projected to PPJ
- **THEN** the theme contains `accentTransforms.accent1.gray: true`
- **AND** the native reference contains `setThemeAccent1Gray` authorized for
  `accentTransforms.accent1.gray`

#### Scenario: Keep unsupported gray topology source-owned

- **GIVEN** accent1 has no gray child, an alpha or another transform, extra
  descendants, duplicate owners, or a non-canonical color owner
- **WHEN** the presentation is projected
- **THEN** the source-bound gray field and capability are omitted

### Requirement: Remove only an existing accent1 gray child

The source-preserving compiler MUST accept one authorized change to the
existing `design.theme.accentTransforms.accent1.gray` field. `true` MUST be a
no-op, while `false` MUST remove only the existing direct `a:gray` child in the
canonical ThemePart. The compiler MUST NOT create a missing gray child.

#### Scenario: Remove and reproject the direct gray child

- **GIVEN** a projected direct accent1 gray owner
- **WHEN** an authorized PPJ program changes `gray` to `false`
- **THEN** a no-op compile preserves bytes and reports no changed parts
- **AND** the false edit changes only the ThemePart and removes `a:gray`
- **AND** a fresh projection omits the accent1 gray transform

#### Scenario: Fail closed for unsafe gray edits

- **GIVEN** a source-bound accent1 gray owner
- **WHEN** the field is deleted, combined with another theme field, given an
  unsupported sibling or invalid type, or its capability is altered
- **THEN** compilation fails without creating a new native gray node
