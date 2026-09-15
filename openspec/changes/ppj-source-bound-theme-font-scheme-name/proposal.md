## Why

F-15 already exposes the six bounded theme font slots, but an imported font scheme's a:fontScheme/@name has no independent PPJ field or capability. This leaves a small, observable ThemePart metadata gap and prevents an agent from changing that one value while preserving the rest of the source package.

## What Changes

- Add optional design.theme.fontScheme.name to the PPJ schema and artifact wire model.
- Project an existing a:fontScheme/@name with a field-qualified setThemeFontSchemeName capability.
- Allow source-free authoring to write the bounded name and allow source-bound edits to change only the existing ThemePart attribute.
- Reject a missing or empty source name, deletion, combined font-slot edits, and capability tampering.
- Add the registry, generated/reference documentation, backlog evidence, and a focused source-bound regression.

## Capabilities

### New Capabilities

- presentation-theme-font-scheme-name: A strict source-bound PPJ field for an existing OOXML font scheme name.

### Modified Capabilities

None.

## Impact

The PPJ schema, protobuf contract, projector/compiler, PPTX theme codec, capability registry, generated/reference documentation, backlog, and NativeAOT codec tests are affected. The protobuf field is additive; no Office wire version change or new dependency is needed. Unsupported theme topology remains source-owned.
