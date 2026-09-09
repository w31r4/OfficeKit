## Purpose

Complete the source-bound lifecycle of PPJ connectorType while preserving endpoint identities, style and unrelated presentation package contents.

## ADDED Requirements

### Requirement: Connector type can be edited under explicit authority

Recognized editable connectors SHALL expose setConnectorType for connectorType. Source-bound edits SHALL support straight, elbow and curved canonical native types, preserving endpoints, arrows, bindings and non-target ZIP members. No-op SHALL preserve original bytes. Unsupported imported geometry SHALL remain source-owned and receive no typed authority.

#### Scenario: Change a connector route family
- **WHEN** a fresh source request changes a connector to another supported connectorType
- **THEN** export SHALL use that native geometry family and fresh projection SHALL recover the requested type while retaining the endpoints and bindings

### Requirement: Invalid type changes fail closed

The compiler MUST reject unsupported or missing connectorType and changed nativeRef authority without producing a candidate file.

#### Scenario: Invalid type
- **WHEN** a source request sets connectorType to an unsupported string
- **THEN** validation SHALL reject it and return no candidate file
