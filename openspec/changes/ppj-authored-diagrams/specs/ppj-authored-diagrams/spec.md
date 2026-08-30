# PPJ authored diagrams

## ADDED Requirements

### Requirement: PPJ compiles finite semantic diagrams to editable native objects

PPJ SHALL lower each supported authored diagram layout to a deterministic group
of ordinary editable Presentation shapes, connectors, text, and images.

#### Scenario: Authored process diagram

- **WHEN** a valid source-free PPJ declares a styled authored process diagram
- **THEN** NativeAOT SHALL compile ordered editable stage shapes and connectors
- **AND** a second import SHALL preserve their stable object IDs and visible text.

#### Scenario: Authored picture diagram

- **WHEN** a valid picture diagram supplies one local content-addressed image
  asset for every node
- **THEN** NativeAOT SHALL compile editable image and label children without
  remote fetching or rasterizing the complete diagram.

### Requirement: Authored diagram layout terminates before native writing

Authored diagram expansion SHALL be finite, bounded, and validated before PPTX
parts are written.

#### Scenario: Invalid authored graph

- **WHEN** a diagram exceeds the node budget, references an unknown parent or
  style, contains a parent cycle, omits a required picture asset, or requests a
  layout whose topology is not satisfied
- **THEN** validation SHALL reject the PPJ before output changes.

### Requirement: Third-party SmartArt remains source-bound

Authored diagram lowering SHALL NOT broaden mutation authority over imported
SmartArt parts.

#### Scenario: Imported SmartArt is unchanged

- **WHEN** a third-party SmartArt element is projected and rebuilt without an
  issued mutation
- **THEN** its original Diagram parts and relationships SHALL remain unchanged.

#### Scenario: Imported topology is changed

- **WHEN** an Agent changes source-bound SmartArt layout, topology, or node IDs
  without a matching capability
- **THEN** the build SHALL fail closed.

