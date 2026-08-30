## Why

OfficeKit's Presentation codec and source-bound import surface now model far more
state than the PPJ Agent reference exposes. The divergence makes a mature native
capability look absent, forces Agents to guess field names, and allows new codec
primitives to land without an authored or imported PPJ route.

## What Changes

- Integrate the stable `pptx-import-primitives-next` work with the PPJ 2.0
  compiler without restoring the retired public MJS/Compose authoring route.
- Replace method-only capability classification with a field-level manifest that
  records PPJ paths, authored compiler support, imported projection/edit support,
  value domains, review ownership, and intentional internal or host-only status.
- Extend PPJ projection, native references, diff lowering, and authored
  compilation for the newly modeled text, paragraph, line, geometry, and style
  primitives where the codec already has a bounded semantic contract.
- Generate a complete progressive PPJ language reference: a short navigation
  page plus exhaustive field/type/value tables and imported capability tables.
- Reject orphan capabilities when a Presentation runtime or codec primitive has
  no PPJ mapping, explicit internal classification, or host-only classification.
- Preserve fail-closed source editing: broader discoverability does not expose
  raw OOXML, XPath, relationship IDs, arbitrary attributes, or unsupported
  topology.

## Capabilities

### New Capabilities

- `presentation-program-capability-parity`: Defines the field-level contract that
  keeps Presentation codec capabilities, PPJ schema, native projection/lowering,
  generated Agent references, review ownership, and maintenance checks aligned.

### Modified Capabilities

None. The completed PPJ 2.0 change remains historical; this change adds a
versioned parity contract without reopening its finished tasks.

## Impact

- Affected code: `src/ppj/`, Presentation protobuf mappings, NativeAOT PPJ
  validator/projector/compiler/lowerer, source-bound Presentation codecs, Help
  generation, Presentations Skill references, and narrow Presentation/PPJ tests.
- Public surface: `.ppj` remains strict JSON with schema
  `office-kit/ppj/v1`; additive optional fields and capability metadata remain
  backward compatible within v1.
- Source preservation: third-party PPTX no-op behavior and unknown native graphs
  remain unchanged and fail closed.
- Delivery: no new runtime, authoring engine, raw patch interface, or Office wire
  version is introduced.
