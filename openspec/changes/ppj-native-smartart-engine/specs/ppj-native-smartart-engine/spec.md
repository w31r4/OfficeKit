# PPJ native SmartArt engine

## ADDED Requirements

### Requirement: Authored PPJ SmartArt is native SmartArt

OfficeKit SHALL compile an authored PPJ SmartArt element to one native diagram
graphic frame with data, layout, style, colors, and cached drawing parts.

#### Scenario: Authored process diagram

- **WHEN** a valid source-free PPJ declares an authored process SmartArt
- **THEN** the PPTX SHALL contain one native diagram frame and its five-part
  dependency closure
- **AND** import without the embedded PPJ SHALL recover typed SmartArt nodes and
  connections rather than an ordinary group.

### Requirement: PPJ keeps a compact semantic facade

PPJ SHALL represent instance content as nodes and connections and SHALL refer to
custom definitions through a content-addressed JSON asset.

#### Scenario: Shared definition edit

- **WHEN** one of two diagrams sharing a definition is edited
- **THEN** the editor SHALL create a new definition asset and repoint only that
  diagram.

### Requirement: Unsupported SmartArt fails closed by affected region

OfficeKit SHALL preserve unknown native regions and SHALL reject a mutation that
requires interpreting an unsupported operator, extension, or relationship.

#### Scenario: Safe text edit beside an unknown layout extension

- **WHEN** an imported diagram has a proven text leaf and an unknown untouched
  layout extension
- **THEN** the text edit MAY proceed without rewriting the layout extension.

#### Scenario: Unsupported layout edit

- **WHEN** a requested definition change touches an unsupported layout node
- **THEN** validation SHALL identify the definition path and produce no PPTX.

### Requirement: Shape detachment is explicit and lossy

OfficeKit SHALL convert SmartArt to ordinary shapes only through an explicit
detach operation over a verified cached drawing.

#### Scenario: Caller detaches a supported cache

- **WHEN** the caller invokes `detachToShapes()` on a diagram with a supported
  cached drawing
- **THEN** the PPJ element SHALL become a group and the result SHALL report that
  SmartArt semantics were removed.
