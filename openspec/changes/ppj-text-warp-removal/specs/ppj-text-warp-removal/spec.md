## Purpose

Make source-bound PPJ preset text warp and its literal adjustment list editable through their full presence lifecycle without losing unrelated presentation state.

## ADDED Requirements

### Requirement: Preset and guide deletion lifecycle
The codec SHALL support removal and restoration of canonical preset text warp for authorized text, shape, master/layout placeholder and table body styles, preserving unrelated XML and ZIP entries.

#### Scenario: Clear guide overrides
- **WHEN** the preset remains and its projected adjustment field is removed or set to an empty array
- **THEN** its direct guides are removed, fresh projection omits the list, and preset and other body properties remain intact.

#### Scenario: Remove and restore the warp
- **WHEN** a projected preset and its dependent guides are removed and later restored
- **THEN** the canonical warp child is removed and restored, explicit textNoShape remains distinguishable from absence, and ordered zero/signed guide values reproject unchanged.

#### Scenario: Remove the whole body style
- **WHEN** an authorized shape or table cell removes a style containing only the supported direct body properties including warp fields
- **THEN** those direct properties are removed and can be restored, including compact table text normalization.

### Requirement: Safe and unambiguous deletion
The codec SHALL preserve existing setter/guide wire fields, reject false or conflicting deletion intent, reject nonempty guides without a preset, and reject duplicate or noncanonical native warp subgraphs.

#### Scenario: Invalid request
- **WHEN** deletion is false, coexists with a preset or guides, or targets an unsupported native warp subgraph
- **THEN** the edit is rejected rather than exporting an approximation.
