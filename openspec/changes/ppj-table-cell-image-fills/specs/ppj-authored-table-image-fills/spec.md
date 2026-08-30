## ADDED Requirements

### Requirement: PPJ table cells compile image fills
The authored PPJ compiler SHALL accept a valid shared image fill on an explicit table cell or inherited `defaultCellFill`, resolve its declared local asset, and compile it as a native DrawingML table-cell image fill without rasterizing the table.

#### Scenario: Explicit cell image fill
- **WHEN** a PPJ table cell declares an image fill with a valid local asset and bounded image-paint options
- **THEN** the output PPTX contains an editable table cell whose `a:tcPr` owns a native `a:blipFill` referencing that asset

#### Scenario: Inherited default image fill
- **WHEN** a PPJ table style declares an image `defaultCellFill` and a cell does not override it
- **THEN** the compiler lowers the inherited image against that cell's computed dimensions and emits the corresponding native image fill

### Requirement: Table image fills preserve PPJ image semantics
Table-cell image fills SHALL use the same asset, fit, crop, tile, opacity, hash, and path validation semantics as other PPJ image paint.

#### Scenario: Invalid asset reference
- **WHEN** a table-cell image fill refers to a missing or hash-invalid asset
- **THEN** PPJ check or build fails before package mutation with a path-specific diagnostic

#### Scenario: Authored recovery
- **WHEN** a PPTX containing authored table-cell image fills is re-imported with its valid embedded PPJ program
- **THEN** OfficeKit restores the exact PPJ fill declarations and stable element identities

### Requirement: Imported unsupported table fills remain source-preserved
The new authored capability MUST NOT reinterpret arbitrary third-party table-cell image topology as safe source-bound PPJ state.

#### Scenario: Unsupported imported table image fill
- **WHEN** a third-party PPTX table contains an image fill outside the bounded imported table profile
- **THEN** OfficeKit preserves it through the source graph or classifies the object as opaque and does not silently rebuild the table
