## Why

PPJ now documents the full Presentation state and imported edit vocabulary, but
its source-free compiler still owns less visual state than the language accepts.
The largest gap with PPTD is not element count: PPJ has more element kinds and a
stronger source-bound model. PPTD devotes most of its surface to fills, charts,
tables, and style inheritance, while several equivalent PPJ fields currently
fail late or are silently ignored.

## What Changes

- Make schema-valid authored visual fields either compile or fail before output;
  no accepted styling field may be silently dropped.
- Lower PPJ custom path geometry through the existing bounded native custom-
  geometry codec.
- Add canonical gradient fill and line-alpha writer state for source-free shapes,
  connectors, and slide backgrounds without weakening imported-source opacity.
- Make the existing chart and table style fields operational where the native
  compiler already has a bounded semantic owner; explicitly reject the rest.
- Generate authored availability from the capability registry and keep the PPJ
  manual synchronized with each new primitive.

## Capabilities

### New Capabilities

- `ppj-authored-visual-primitives`: Defines native source-free compilation and
  fail-closed behavior for PPJ geometry, fill, stroke, chart, and table state.

## Impact

- Affected code: PPJ schema/validator/compiler/projector, additive Presentation
  wire fields, native PPTX shape/background/line/chart/table writers, capability
  registry, generated PPJ reference, and one existing integrated PPJ test.
- Wire protocol remains version 2; additions are optional protobuf fields.
- Imported PPTX no-op and opaque preservation remain unchanged.
- This change does not copy PPTD syntax or implementation. PPTD documentation is
  used only to identify public authoring categories for clean-room parity work.
