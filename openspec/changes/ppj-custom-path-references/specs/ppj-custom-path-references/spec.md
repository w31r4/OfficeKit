## Purpose

Preserve native reference identity in every supported PPJ custom-shape path coordinate and arc parameter through source editing.

## ADDED Requirements

### Requirement: Custom path parameters retain native references

Custom-shape moveTo/lineTo/quadraticTo/cubicTo point slots and arcTo radiusX/radiusY/startAngle/sweepAngle SHALL accept existing numeric values or native built-in/declared adjustment/guide reference strings. Numbers SHALL retain existing viewBox and degree conversions; references SHALL retain native identity and units without conversion. Export and fresh projection SHALL preserve mixed parameters. Unresolved references and invalid resolved arc parameters MUST reject. Standalone line and mask/clip referenced paths MUST remain rejected.

#### Scenario: Referenced curves and arc
- **WHEN** a custom shape uses mixed literal/reference points, controls and arc parameters
- **THEN** fresh projection SHALL retain all references and numeric values under existing normalization

### Requirement: Source path edits preserve the graph

Source paths authority SHALL permit reference/literal replacements on supported common-positive-viewport custom shapes. Source no-op SHALL preserve bytes. Edits SHALL preserve unchanged graph controls, other path slots, text, frame and non-target ZIP members. Adjustment/guide edits SHALL retain dependent path references and validate their resolved values.

#### Scenario: Edit references and dependencies
- **WHEN** original-source requests change path references or a referenced adjustment formula
- **THEN** requested state SHALL reproject and dependency evaluation SHALL reflect the changed formula without flattening path references

### Requirement: Preview reports reference limitations

Preview SHALL explicitly diagnose custom path references whose complete meaning it cannot render.

#### Scenario: Referenced path preview
- **WHEN** a path contains a native reference
- **THEN** unrendered parameter semantics SHALL NOT be reported as fully supported
