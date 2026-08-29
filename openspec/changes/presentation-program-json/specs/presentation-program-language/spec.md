## ADDED Requirements

### Requirement: Strict single-file PPJ program
The system SHALL accept a `.ppj` file only as UTF-8 strict JSON with schema identifier `office-kit/ppj/v1`, SHALL reject unknown fields and executable content, and SHALL enforce the declared document and expansion budgets before native output changes.

#### Scenario: Valid PPJ is accepted
- **WHEN** a program uses the v1 schema, typed fields, unique stable IDs, local assets, and remains within all budgets
- **THEN** the system validates the complete program and returns its canonical SHA-256

#### Scenario: Executable or unbounded content is rejected
- **WHEN** a program contains functions, recursion, `while`, raw OOXML, arbitrary expressions, unknown fields, or exceeds an expansion budget
- **THEN** validation fails with a path-specific error before compilation

### Requirement: Typed presentation state
The system SHALL model authored persistent state with typed deck, page, text, shape, image, chart, table, connector, group, media, placeholder, native, and opaque structures instead of a generic property bag.

#### Scenario: Type-specific validation
- **WHEN** a text element contains a chart-only field or an image element omits its source
- **THEN** validation identifies the element ID, field, expected type, and invalid value

### Requirement: Ordered pages and scene elements
The system SHALL treat page order and each page or group's element array order as semantic ordering and native z-order.

#### Scenario: Background image, overlay, and text retain order
- **WHEN** PPJ declares an image, a translucent overlay, and text in that sequence
- **THEN** the compiled PPTX places the image behind the overlay and editable text

### Requirement: Simple and rich text
The system SHALL accept plain strings for ordinary text and structured paragraphs/runs for mixed formatting without requiring HTML or executable markup.

#### Scenario: Mixed formatting round trip
- **WHEN** a text element contains multiple paragraphs and differently styled runs
- **THEN** build and reimport preserve the text, paragraph order, run styles, language, and stable element ID

### Requirement: Finite component templates
The system SHALL expand components with typed parameters, slots, variants, bounded finite repeats, simple conditions, local coordinates, and explicit frames; it SHALL reject recursive or ambiguous expansion.

#### Scenario: Finite repeated evidence rows
- **WHEN** a component repeats over a finite keyed evidence array within the configured limits
- **THEN** the compiler produces deterministic stable instance IDs and identical output for identical input

#### Scenario: Component cycle is rejected
- **WHEN** component A eventually instantiates component A or creates duplicate expanded IDs
- **THEN** validation fails before any PPTX bytes are written

### Requirement: Local immutable assets
The system SHALL resolve assets only from relative local paths bound to MIME and SHA-256 and SHALL reject network fetches, directory escapes, missing files, and stale bytes.

#### Scenario: Matching local asset compiles
- **WHEN** the declared relative file exists and its bytes, MIME, and hash match
- **THEN** the compiler may embed or reuse it and records the asset mapping in its receipt

### Requirement: Deterministic authored compilation
The NativeAOT C# compiler SHALL compile a validated source-free PPJ into a native editable PPTX and return a receipt containing program hash, output hash, stable mapping, and package footprint.

#### Scenario: Repeated build is deterministic
- **WHEN** the same PPJ and asset bytes are built repeatedly from clean output paths
- **THEN** canonical package hashes, stable mappings, and visible output are identical

### Requirement: Embedded authored program
An OfficeKit-authored PPTX SHALL contain the canonical PPJ snapshot and stable node/asset mapping as reserved OPC parts.

#### Scenario: Authored PPTX contains recoverable program
- **WHEN** a source-free PPJ is built successfully
- **THEN** reimport can recover its program, design intent, assets, page IDs, element IDs, and native mappings without reconstructing them heuristically
