## Why

OfficeKit can already compose visually free presentations, but the shipped authoring path does not connect delivery intent, composition intent, native animation, chart builds, Morph, and review into one coherent Agent workflow. The result is either a static deck or isolated effects applied after design, instead of motion that advances the argument and survives round-trip editing.

## What Changes

- Extend the durable presentation authoring plan with delivery mode, motion policy, page composition intent, and bounded motion intent without adding a second plan format.
- Add typed slide animation and Morph operations for fade, wipe, fly, zoom, pulse, text builds, chart builds, trigger order, and explicit object pairing.
- Compile the typed operations into a bounded canonical PresentationML timing graph while preserving unknown imported timing and Morph structures as opaque source content.
- Make the C authoring route select motion from `reader`, `hybrid`, or `live` delivery intent and expose six reusable communication recipes.
- Add motion inspection and review evidence covering targets, order, chart builds, Morph pairs, plan consistency, over-animation, and honest native-host playback status.
- Validate the complete workflow with three focused real presentation scenarios rather than expanding the general benchmark matrix.

## Capabilities

### New Capabilities

- `presentation-motion-authoring`: Durable motion intent, delivery-aware selection, composition coupling, and six communication-oriented motion recipes.
- `presentation-motion-runtime`: Typed JavaScript operations, additive wire messages, canonical native timing/Morph compilation, inspection, limits, and opaque imported preservation.
- `presentation-motion-review`: Structural motion review, plan consistency, over-animation warnings, playback evidence levels, and focused real-world acceptance.

### Modified Capabilities


## Impact

- Public Presentation model and Help catalog under `src/presentation/`, `src/help/`, and the package root exports.
- Presentation authoring plan validation, task descriptors, REPL resume state, and review reports.
- Additive protocol messages in `proto/office_kit/artifact/v1/office_artifact.proto`; Office wire version remains 2.
- C# canonical PPTX timing and Morph writer plus generated audited WASM runtime.
- Presentations Skill, recipes, API documentation, coverage, focused tests, and three dogfood artifacts.
- No PDF, Spreadsheet, Document, Live Add-in, provider, or template asset format changes.
