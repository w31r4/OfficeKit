## ADDED Requirements

### Requirement: Complete PPJ field ownership
Every accepted public PPJ field SHALL have a field-level capability record that
identifies its schema path, persistent-state meaning, authored compiler owner,
projection owner, review owner, or explicit metadata-only classification.

#### Scenario: Accepted authored field has an owner
- **WHEN** the schema accepts a Presentation field that can change native output
- **THEN** the parity check finds its authored compiler and review ownership and the compiler does not silently ignore it

#### Scenario: Unowned field fails maintenance
- **WHEN** a new schema property is added without an operational owner or an explicit metadata-only classification
- **THEN** the capability parity check fails with the unresolved schema path

### Requirement: Closed imported native leaves
An imported PPJ native reference SHALL expose bounded editable leaves only as a
closed kind, opaque stable leaf ID, typed scalar value, and expected hash, and
SHALL NOT expose package locators or arbitrary attribute names.

#### Scenario: Issued native leaf is editable
- **WHEN** a projected PPTX object contains a native leaf whose source topology and value domain are proven by the codec
- **THEN** PPJ contains the leaf kind, current value, opaque leaf ID, and expected hash under that object's native reference

#### Scenario: Unknown native property remains opaque
- **WHEN** an imported property is not represented by a registered bounded leaf kind
- **THEN** the property stays in the source package and PPJ does not expose a writable generic property

### Requirement: Source-bound leaf lowering
The PPJ compiler SHALL reproject the exact source and lower only changed,
current native leaves through the existing source-bound edit-plan codec.

#### Scenario: Multiple bounded leaf edits compile locally
- **WHEN** an Agent changes supported scalar leaf values with current source, object, capability, leaf, and expected hashes
- **THEN** the compiler applies only those leaf operations, reports the mutation footprint, and preserves unrelated package content

#### Scenario: Stale leaf fails before output
- **WHEN** a leaf ID, kind, expected hash, source revision, or enclosing object hash is stale or mismatched
- **THEN** build fails closed and writes no output artifact

### Requirement: Authored and imported semantic parity
For a capability that has a complete authored semantic model, PPJ SHALL use the
same typed value vocabulary for source-free creation and imported projection;
partial source-only semantics SHALL remain nativeRef-only.

#### Scenario: Complete text and line semantics share vocabulary
- **WHEN** font, paragraph, fill, stroke, geometry, or body-layout semantics are complete in both paths
- **THEN** the authored field, projected field, compiler mapping, and generated reference use the same names, units, and value domain

#### Scenario: Partial native semantics do not masquerade as authored state
- **WHEN** only one direct OOXML token is safely editable but inherited or effect-bearing semantics are incomplete
- **THEN** the capability is documented and represented as a native leaf rather than a general authored style field

### Requirement: Exhaustive generated PPJ reference
The generated `ppj.md` SHALL document every root field, typed element, nested
definition, property, enum or scalar constraint, and registered native leaf kind
from the schema and capability registry.

#### Scenario: Agent searches for a detailed primitive
- **WHEN** an Agent searches `ppj.md` for a registered field or native leaf such as line join, text-body inset, capitalization, or paragraph spacing
- **THEN** the reference returns its PPJ location, accepted values or units, authored/imported availability, and failure boundary

### Requirement: Primitive update gate
A Presentation runtime, codec, schema, Help, or review change SHALL update its
PPJ ownership record or explicitly classify the capability as internal,
inspection-only, or host-only before the maintenance gate passes.

#### Scenario: Codec primitive lands without PPJ mapping
- **WHEN** a new public Presentation field or native leaf kind appears in the codec or wire contract without a matching registry entry and reference route
- **THEN** the maintainer check fails and names the missing capability

#### Scenario: Host-only operation remains outside PPJ
- **WHEN** a capability operates only on an open desktop PowerPoint session
- **THEN** the registry classifies it as host-only and the PPJ schema does not serialize it
