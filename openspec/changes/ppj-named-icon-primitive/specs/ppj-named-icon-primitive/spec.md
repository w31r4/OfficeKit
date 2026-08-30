## ADDED Requirements

### Requirement: Offline named icon element
PPJ SHALL accept a finite `icon` element whose `iconName` resolves only against
the pinned compiler catalog and SHALL compile it without network access or an
external asset declaration.

#### Scenario: Agent authors a known free icon
- **WHEN** an icon uses a catalog name such as `fas:lightbulb`
- **THEN** build emits editable native DrawingML geometry with the declared
  frame, paint, transform, accessibility and z-order

#### Scenario: Icon name is unknown
- **WHEN** `iconName` does not exist in the pinned catalog
- **THEN** semantic validation rejects before native presentation writing

### Requirement: Deterministic icon recovery
OfficeKit-authored presentations SHALL recover the exact named-icon PPJ state
from the embedded program snapshot.

#### Scenario: Authored icon is rebuilt and reimported
- **WHEN** the PPTX still contains the matching OfficeKit program and node map
- **THEN** reimport returns the same stable element ID and `iconName`

#### Scenario: Imported geometry has no OfficeKit program
- **WHEN** a third-party PPTX contains visually similar custom geometry
- **THEN** projection classifies it as ordinary geometry and does not infer a
  catalog icon identity

### Requirement: Catalog provenance
The distributed catalog SHALL identify its exact source version and carry the
required Font Awesome Free license notices.

#### Scenario: Release inventory is inspected
- **WHEN** the package contents and third-party notices are reviewed
- **THEN** the generated catalog version, CC BY 4.0 icon license and applicable
  MIT code license are discoverable without fetching the npm packages
