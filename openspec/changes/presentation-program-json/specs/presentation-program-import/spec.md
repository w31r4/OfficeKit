## ADDED Requirements

### Requirement: Source-bound PPJ projection
The system SHALL project an arbitrary PPTX into a PPJ whose source points to a read-only content-addressed relative copy of the original package.

#### Scenario: Import creates portable source binding
- **WHEN** a third-party PPTX is imported to `deck.ppj`
- **THEN** the source bytes are stored under `deck.assets/source/<sha256>.pptx` and PPJ records the matching relative URI and hash

### Requirement: Complete visible-object classification
The projection SHALL represent every visible imported object as a typed element or an opaque descriptor with stable identity, page, frame, summary, revision, and available native capabilities.

#### Scenario: Unsupported visible object remains discoverable
- **WHEN** a slide contains a visible object outside the typed model
- **THEN** PPJ includes one opaque element for it and the source package retains its unknown native graph unchanged

### Requirement: Capability-bound native references
An imported `nativeRef` SHALL expose only compiler-issued bounded operations and expected hashes and SHALL never expose raw XML, XPath, arbitrary attributes, part paths, or relationship identities.

#### Scenario: Stale native reference fails closed
- **WHEN** the source hash, object hash, topology, or capability no longer matches the imported evidence
- **THEN** build rejects the affected edit without changing the source or unrelated content

### Requirement: Byte-identical imported no-op
Building an unchanged projected PPJ SHALL return the original source bytes exactly and SHALL NOT add embedded PPJ parts to a third-party package.

#### Scenario: Third-party no-op is identical
- **WHEN** PPJ has no semantic difference from a fresh projection of its bound source
- **THEN** output SHA-256 and bytes equal the source and no new OPC part or relationship is introduced

### Requirement: Local lossless edit lowering
The compiler SHALL compare edited PPJ with a fresh source projection, lower supported changes into typed Edit Plan operations, re-prove all preconditions, and preserve every non-target source structure.

#### Scenario: Supported native leaf changes locally
- **WHEN** an Agent changes a capability-issued native text leaf with a current expected hash
- **THEN** only the target XML token and required package metadata change, while unrelated parts and opaque graphs remain identical

#### Scenario: Unsupported opaque mutation is rejected
- **WHEN** an Agent changes an opaque field without an issued capability
- **THEN** build reports the element and blocked field and produces no output file

### Requirement: Authored program recovery authority
When an imported PPTX contains a valid OfficeKit embedded PPJ, the system SHALL restore that PPJ as authoritative; if the native presentation drifted while the embedded program remained, the system SHALL not merge the native drift.

#### Scenario: Exact embedded recovery
- **WHEN** embedded program and native fingerprints match
- **THEN** import restores the exact PPJ bytes and stable mapping

#### Scenario: Native drift is ignored without overwriting input
- **WHEN** an external editor changes an authored PPTX but retains its embedded PPJ
- **THEN** import restores the embedded PPJ, does not prompt or merge the drift, and any later build writes only to a distinct output path

### Requirement: Missing embedded program falls back to projection
The system SHALL use ordinary third-party projection when the reserved authored-program parts are absent or structurally unreadable.

#### Scenario: External editor strips OfficeKit parts
- **WHEN** a PPTX no longer contains a usable embedded program and mapping
- **THEN** import creates a source-bound projected PPJ and does not claim exact program recovery
