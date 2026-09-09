## Purpose

Preserve explicit depth and absence when editing PPJ flat-text body styles backed by an original PowerPoint package.

## ADDED Requirements

### Requirement: Flat-text depth presence lifecycle
The codec SHALL distinguish absent flatTextZ from explicit signed 32-bit values, including zero, for text, shape, master/layout placeholder and table cell body styles with existing editing authority.

#### Scenario: Remove and restore depth
- **WHEN** an authorized source-bound style removes flatTextZ and later restores a bounded value
- **THEN** the canonical flatTx child is removed and restored, fresh projection reflects each state, and unrelated source XML and ZIP entries remain unchanged.

#### Scenario: Remove a simple style
- **WHEN** a shape or table cell removes a style consisting of removable body properties including flatTextZ
- **THEN** direct depth is removed and can be restored through the existing style authority, including compact table text normalization.

### Requirement: Unambiguous safe deletion
The wire SHALL retain setter field 39 and use additive true-only deletion field 45 without a coexisting setter. Noncanonical or duplicate flatTx content SHALL be rejected for deletion.

#### Scenario: Invalid edit intent or source
- **WHEN** deletion is false, coexists with a setter, or targets duplicate or noncanonical flatTx content
- **THEN** the codec rejects the edit without silently discarding source content.
