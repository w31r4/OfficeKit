## Purpose

Preserve custom path extrusion eligibility as an optional PPJ boolean across authored and source-bound presentation edits.

## ADDED Requirements

### Requirement: Extrusion eligibility preserves optional presence

Custom paths SHALL accept optional boolean extrusionOk. Authored output and fresh projection SHALL preserve true, false and omission independently. Non-boolean values MUST reject. This attribute SHALL NOT imply complete 3-D rendering support.

#### Scenario: Three distinct states
- **WHEN** a path has true, false or omitted extrusionOk
- **THEN** export and fresh projection SHALL recover the same presence and value

### Requirement: Source path edits retain extrusion eligibility

Existing path-edit authority SHALL allow adding, changing and removing extrusionOk on supported custom geometry. Source no-op SHALL preserve bytes. Editing other path fields SHALL preserve extrusionOk; editing extrusionOk SHALL preserve commands, fill/stroke, text, frame and non-target package members.

#### Scenario: Original-source edits
- **WHEN** separate requests modify or remove extrusionOk or modify another path field
- **THEN** fresh projection SHALL recover the requested field state and preserve unrelated data

### Requirement: Preview reports unsupported extrusion

Preview SHALL retain explicit limitations for extrusion eligibility when it cannot render the corresponding 3-D semantics.

#### Scenario: Preview with extrusion eligibility
- **WHEN** a path carries extrusionOk
- **THEN** the field SHALL NOT be reported as fully rendered
