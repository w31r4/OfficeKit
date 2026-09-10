## Purpose

Let PPJ change the direct starting value of an existing numbered paragraph while preserving the original list scheme, styling and all unrelated presentation content.

## ADDED Requirements

### Requirement: Numbered paragraph start value lifecycle

PPJ SHALL accept integer `bullet.startAt` values from 1 through 32767 on numbered bullets. Existing source-bound text and shape paragraphs SHALL support adding, replacing, deleting and restoring that direct field using its exact `setTextParagraphStyle` capability. Deletion SHALL remove only the direct start attribute and SHALL preserve the numbered marker and scheme.

#### Scenario: Change, remove and restore a direct start value

- **WHEN** a projected numbered paragraph changes its start value, removes it, then restores it from each actual candidate and fresh projection
- **THEN** the native attribute and reprojected PPJ SHALL reflect each request, including explicit 1 versus absence, while adjacent paragraphs, list scheme/style, runs and non-target package content remain unchanged

#### Scenario: Invalid request or missing authority

- **WHEN** a start value is null, fractional, a string, zero or outside the declared range, or the exact field capability is absent
- **THEN** compilation SHALL reject the request without emitting a candidate

### Requirement: Preserve unmodeled numbered source content

Malformed, duplicate or unsupported native markers SHALL remain source-owned. No-op compilation SHALL preserve source bytes; unrelated paragraph edits SHALL preserve unknown marker XML and unchanged lexical attribute values. A start-only request SHALL NOT change marker kind, scheme, asset relationships or paragraph topology.

#### Scenario: Unsupported or malformed native marker

- **WHEN** a source has an invalid start token, unknown numbering scheme or ambiguous marker choice
- **THEN** projection SHALL avoid throwing or inventing a modeled value, unrelated edits SHALL preserve that source content, and replacement SHALL fail closed

### Requirement: Preview evidence remains explicit

The public field reference SHALL document the direct value and source edit boundaries. Preview SHALL diagnose automatic numbering as partial or unavailable wherever it cannot paint the numbering state.

#### Scenario: Start value in an automatic list

- **WHEN** preview receives a numbered paragraph with an explicit start value
- **THEN** unsupported numbering SHALL remain visible in diagnostics without claiming complete list layout
