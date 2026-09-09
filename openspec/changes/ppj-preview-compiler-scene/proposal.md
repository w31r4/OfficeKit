## Why

G-01 remains open: local preview draws canonical PPJ while the PPTX writer consumes expanded, catalog-resolved native slides. This duplicates component, dataset and chart interpretation and makes preview disagree with the same program's export. G-11 now reports these limitations; the next step is to share the actual compiler output.

## What Changes

- Add an opt-in, versioned preview-scene result to PPJ compilation, reusing the existing `PresentationArtifact` writer IR rather than introducing another authoring language or a JavaScript semantic compiler.
- Capture source-free slides from the same `MaterializeSlide` invocation used by the PPTX writer, including compiled component geometry, catalog styles, normalized chart data and already-lowered vector chart groups.
- Build source-bound scenes from the actual resulting candidate through the existing C# codec, including precise native-leaf edits and no-op reuse. Preserve opaque boundaries, source identity, semantic-to-native ownership and original input bytes.
- Carry page/node provenance, native asset identity, scene version/hash and candidate hash. Preserve canonical `programJson` and its existing public meaning.
- Make local preview request and consume this scene through a leaf adapter. Remove component/chart layout guesses from the canonical-PPJ drawing route; retain explicit limitations for native features the SVG painter cannot yet draw.
- Keep G-11 support/reliability and G-12 publication protection. Only retire a previous limitation when a mapped scene state and regression prove it resolved; missing/incompatible scene evidence fails closed.
- Verify high-level and equivalent explicit authored inputs produce equivalent scene geometry, data and resolved styles, and verify the actual SVG consumer uses those values. Exercise source-bound edit/reprojection, both native codec entrypoints and the rebuilt runtime.

This closes the shared-input boundary, not every painter gap G-02–G-10. Preserving full native fields with explicit drawing limits is required; truncating the scene to easy rectangles or replacing all complex content with opaque placeholders does not satisfy G-01.

## Capabilities

### New Capabilities

- `ppj-preview-compiler-scene`: opt-in compiler-owned preview scenes, actual-candidate provenance, transport and local-preview consumption.

### Modified Capabilities

None. No main spec currently defines this scene. The existing G-11 diagnostic and G-12 publication contracts remain required and are not weakened.

## Impact

Affected areas: `proto/office_kit/artifact/v1/office_artifact.proto`, generated bindings, `PpjAuthoredPresentationCompiler`, `PpjPresentationCompiler`, both `PpjCodecProtocol` and `CodecProtocol`, PPJ native-host build/packaging checks, `src/ppj/native.mjs`, `workspace.mjs`, preview/assessment/publication modules, registry and coverage metadata, focused native/JS tests, output documentation and presentation review guidance. No new graphics dependency, Office installation, network service or public JavaScript Presentation object model is introduced. Ordinary build/check calls remain scene-free by default. Concurrent trendline-label protocol/compiler work must be preserved.
