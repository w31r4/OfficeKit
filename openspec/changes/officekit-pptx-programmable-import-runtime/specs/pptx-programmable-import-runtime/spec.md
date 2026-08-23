## ADDED Requirements

### Requirement: Every visible imported object has one explicit state

For a trusted imported PPTX, `presentation.inspect()` SHALL be able to return
one deterministic source-bound record for every direct visible slide
shape-tree object. Each record SHALL have exactly one of `typed-editable`,
`native-leaf-editable`, `source-derived-reusable`, or `opaque-preserved` as its
primary classification and SHALL list only codec-issued operations.

#### Scenario: Unsupported complex object

- **WHEN** a visible imported object has neither a typed edit, issued native
  leaf, nor source-derived reuse capability
- **THEN** inspection reports it as `opaque-preserved` with a bounded reason
  while unrelated objects remain editable

#### Scenario: Inspection cannot grant authority

- **WHEN** a caller modifies an inspection record or reuses it against another
  source revision
- **THEN** no mutation permission changes and the actual edit API rejects stale
  or unissued authority

### Requirement: Completeness is independently proved

The benchmark SHALL compare direct visible children from the original slide
shape trees with runtime classification locators. A self-reported imported
object count SHALL NOT be sufficient evidence.

#### Scenario: Importer omits one visible source object

- **WHEN** the raw source contains a direct visible shape-tree child without a
  matching classified source locator
- **THEN** the completeness oracle fails and identifies the slide and index

### Requirement: Safe SVG style edits are token-local and source-bound

OfficeKit SHALL issue SVG color, opacity, and bounded transform leaves only for
direct, unambiguous, inactive SVG tokens. Applying a leaf SHALL require the
current source and value hashes and SHALL preserve every unissued byte.

#### Scenario: Direct color edit

- **WHEN** an Agent applies a current issued direct fill or stroke leaf
- **THEN** only the declared token changes and masked SVG plus all non-target
  package parts remain byte-identical

#### Scenario: Unsafe styling topology

- **WHEN** styling depends on CSS, inheritance, a paint server, active content,
  an external resource, or an unissued transform structure
- **THEN** OfficeKit reports no edit capability and preserves the SVG unchanged

### Requirement: Source-derived continuation retains source authority

A reused source slide or component SHALL be re-importable and SHALL expose the
same typed and issued-leaf workflow for supported edits. Unknown source
subgraphs SHALL remain attached and SHALL not be reconstructed as semantic
objects merely to permit continuation.

#### Scenario: Resume after reviewed revision

- **WHEN** a fresh Agent context resumes a task from its latest reviewed PPTX
- **THEN** it re-imports that revision, regenerates source-bound locators, and
  can continue supported edits without restoring a JavaScript heap
