## Purpose

Represent each paragraph's explicit writing direction while preserving source fidelity, independent style edits and a clear boundary around host bidirectional layout.

## ADDED Requirements

### Requirement: Explicit paragraph direction
A typed paragraph style SHALL accept direction as left-to-right or right-to-left, independently of alignment, column order and vertical text. Authored precedence SHALL follow the existing paragraph style layers. Native export and fresh projection SHALL distinguish either explicit direction from absence. Invalid types and unsupported strings SHALL reject.

#### Scenario: Author and recover explicit LTR
- **WHEN** a paragraph explicitly selects left-to-right above a lower-priority right-to-left style
- **THEN** native output SHALL retain explicit false for paragraph RTL and fresh projection SHALL recover left-to-right

### Requirement: Independent source direction lifecycle
Ordinary source-bound text/shape paragraphs with supported inline topology SHALL support add, set, remove and restore of direction under the exact text.paragraphs[].style.direction authority. Removing direction or a direction-only style wrapper SHALL clear only the recognized direct attribute. Each edit SHALL retain unrelated properties, neighboring paragraphs, runs and non-target package content.

#### Scenario: Edit actual source and restore
- **WHEN** a fresh projection is changed from explicit RTL to LTR, then removed and restored using each candidate as source
- **THEN** native attribute presence and PPJ projection SHALL match each requested state while other XML and ZIP content remain

#### Scenario: Missing direction authority
- **WHEN** direction is changed without its exact issued source capability
- **THEN** compilation SHALL reject without a candidate

### Requirement: Source spelling and unknown direction remain intact
Canonical native booleans and their equivalent word spellings SHALL retain their original spelling on no-op and unrelated edits. Unknown direction tokens SHALL remain source-owned, omitted from modeled direction and protected against replacement; independent modeled edits SHALL remain available.

#### Scenario: Unknown native direction
- **WHEN** an imported paragraph has an unrecognized RTL token
- **THEN** unrelated edits SHALL retain it, projection SHALL omit modeled direction and explicit replacement SHALL reject

### Requirement: Discoverability and honest preview
Schema, Help, capability documentation, generated references and Agent guidance SHALL agree on direction values, presence, owner and source lifecycle. Preview SHALL report unresolved bidirectional layout and SHALL NOT claim full support.

#### Scenario: Preview explicit paragraph direction
- **WHEN** a native preview scene contains explicit paragraph direction
- **THEN** its unresolved direction SHALL appear in machine-readable diagnostics with a non-supported status
