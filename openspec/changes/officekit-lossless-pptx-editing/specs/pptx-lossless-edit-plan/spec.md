## ADDED Requirements

### Requirement: Proven no-op preserves the complete source package
OfficeKit SHALL return the exact imported PPTX byte sequence only when the complete semantic projection is unchanged and SHALL otherwise execute an authorized edit path or fail closed.

#### Scenario: Complete no-op
- **WHEN** an imported presentation, including presentation-level and slide-level state, is unchanged
- **THEN** export returns bytes whose SHA-256 and contents equal the source package

#### Scenario: Non-slide state changed
- **WHEN** a custom show, section, view property, comment, note, clone, or other presentation-level state changes
- **THEN** OfficeKit does not classify the export as a no-op

### Requirement: Edit Plans are finite and source-bound
OfficeKit SHALL compile each supported imported edit into a finite plan containing a source revision, operation ID, target binding, old-value preconditions, new value, and compiler-generated mutation footprint.

#### Scenario: Safe text leaf
- **WHEN** exactly one supported text run changes and all other projected state matches the source
- **THEN** the compiler emits a text-leaf operation bound to the source package, SlidePart, native shape-tree index, element hashes, leaf ordinal, and old value

#### Scenario: Stale or ambiguous target
- **WHEN** any source, part, element, semantic, leaf, or old-value precondition is stale or ambiguous
- **THEN** the codec rejects the operation without changing an output package

### Requirement: Accepted edits preserve undeclared content
The PPTX codec SHALL mutate only declared leaves and required dependent content, SHALL preserve every non-target OPC part content byte-for-byte, and SHALL prove masked equality for each changed XML part.

#### Scenario: Single SlidePart text edit
- **WHEN** one source-bound `a:t` leaf is edited
- **THEN** only its SlidePart changes and replacing the new token with the original token recovers the exact source SlidePart bytes

#### Scenario: Dependent ChartPart title edit
- **WHEN** one issued chart-title run belongs to a uniquely bound internal ChartPart
- **THEN** only that ChartPart changes, replacing the new token recovers its exact source bytes, and the owning graphicFrame, relationship, chart data, and plot topology remain unchanged

#### Scenario: Dependent chart-data edit
- **WHEN** one issued direct numeric bar-chart cache point resolves to exactly one cell in a uniquely bound embedded XLSX
- **THEN** only the ChartPart and embedded XLSX change, masking the cache token and nested worksheet token recovers both exact source parts, and every other outer and embedded OPC part remains byte-identical

#### Scenario: Dependent SmartArt run edit
- **WHEN** one issued direct SmartArt text run resolves to one model ID and run index in a canonical closed DiagramDataPart with a unique inbound owner
- **THEN** only that DiagramDataPart changes, masking the new text token recovers its exact source bytes, and node identity, run topology, relationships, layout, quick-style, colors, and the owning graphicFrame remain unchanged

#### Scenario: Unexpected scope expansion
- **WHEN** execution changes any OPC part outside the compiled footprint
- **THEN** the codec rejects the result as a scope violation

### Requirement: Native leaf editing is capability-issued
OfficeKit SHALL expose only revision-bound safe native leaves returned by inspection and SHALL reject arbitrary XML paths, attributes, namespaces, relationships, identities, or topology changes.

#### Scenario: Issued native leaf
- **WHEN** an Agent edits an inspected leaf ID with its expected hash and an allowed value
- **THEN** OfficeKit compiles the bounded leaf mutation and returns its audit footprint

#### Scenario: Raw XML request
- **WHEN** a caller supplies XPath, part path, raw XML, relationship ID, arbitrary attribute name, or an unissued leaf ID
- **THEN** OfficeKit rejects the request without fallback

### Requirement: Unknown unrelated structures do not block a safe target
OfficeKit SHALL evaluate editability at the selected target and SHALL preserve unrelated unknown geometry or native graphs without projecting them through the authoring subset.

#### Scenario: Complex third-party slide
- **WHEN** a supported text leaf is selected on a slide or presentation containing unrelated unknown geometry, groups, charts, SmartArt, OLE, animation, notes, or comments
- **THEN** those unrelated structures do not by themselves block the edit and remain unchanged
