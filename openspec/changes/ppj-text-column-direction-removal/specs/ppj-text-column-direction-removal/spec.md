## Purpose

Complete columnDirection presence semantics for PPJ source-bound text styles while preserving unrelated presentation content and existing edit authority.

## ADDED Requirements

### Requirement: Column direction lifecycle
The system SHALL distinguish left-to-right, right-to-left and omission for existing editable text, shape, master/layout placeholder and table-cell text styles. Removing a projected field SHALL remove native rtlCol. A style containing only columnDirection, upright and/or rotation SHALL be removable under existing capability checks.

#### Scenario: Remove and restore direct direction
- **WHEN** a projected source style loses columnDirection and is later restored to left-to-right or right-to-left
- **THEN** the native attribute and fresh projection SHALL respectively be absent, explicit false/left-to-right, or explicit true/right-to-left.

#### Scenario: Remove the style owner
- **WHEN** the caller removes a style containing only the supported removable properties
- **THEN** those direct properties SHALL be absent and a compact table cell SHALL accept equivalent structured text with a restored direction.

### Requirement: Preservation and honest preview
Source no-op SHALL preserve bytes. Direction edits SHALL preserve other body attributes, frames, text topology and non-target ZIP entries. Missing authority and overbroad style removal SHALL remain rejected. Preview SHALL retain an explicit unsupported-layout diagnostic.

#### Scenario: Isolated source edit
- **WHEN** an authorized caller changes only columnDirection
- **THEN** only the target body attribute SHALL change and unsupported preview layout SHALL remain diagnosed.
