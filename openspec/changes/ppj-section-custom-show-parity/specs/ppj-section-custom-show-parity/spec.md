## ADDED Requirements

### Requirement: Source-bound section editing
A canonical imported section SHALL advertise bounded name and membership
capabilities while retaining fixed section topology and identity.

#### Scenario: Repartition two imported sections
- **WHEN** an Agent renames one capable section and moves a page boundary while
  preserving one complete in-order partition
- **THEN** build changes only `ppt/presentation.xml` and reimport recovers the
  exact names and membership

### Requirement: Source-bound custom-show editing
A canonical imported custom show SHALL advertise bounded name and ordered page
membership capabilities while retaining fixed show topology and identity.

#### Scenario: Rename and reorder one custom show
- **WHEN** an Agent edits a capable custom show's name and ordered page list
- **THEN** build preserves its native identity and reimport recovers the exact
  name, order, and repeated references

### Requirement: Opaque route graphs fail closed
Opaque, extension-bearing, ambiguous, or stale section/custom-show graphs SHALL
remain source-preserved and SHALL NOT advertise mutation capabilities.

#### Scenario: Unsupported native graph
- **WHEN** an Agent requests a section or custom-show mutation without a fresh
  issued capability
- **THEN** build rejects it without reconstructing the native graph

### Requirement: Agent discoverability
The generated PPJ reference SHALL explain the distinct section-partition and
custom-show-subset semantics and their source-bound edit boundary.

#### Scenario: Agent plans alternate delivery routes
- **WHEN** an Agent reads the PPJ reference
- **THEN** it can distinguish presentation order, section membership, hidden
  slides, and custom-show order
