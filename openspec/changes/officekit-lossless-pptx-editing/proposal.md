## Why

OfficeKit can preserve many imported PowerPoint structures, but a small semantic edit may still reserialize an entire SlidePart or be rejected because an unrelated shape uses geometry outside the typed model. Complex third-party presentations need a source-bound editing path whose success is measured against the original package: exact no-op bytes, a declared mutation footprint, and zero drift outside the selected leaf.

## What Changes

- Add a three-layer compiler boundary for imported PPTX: a lossless source graph, the existing semantic `Presentation` projection, and a finite source-bound Edit Plan.
- Return the exact source package for a proven semantic no-op.
- Compile supported single-leaf edits with source revision, slide, element, semantic, and old-value preconditions; apply them by token splice instead of SlidePart reserialization.
- Preserve all non-target OPC part contents byte-for-byte and require masked equality for each changed XML part.
- Add controlled native-leaf inspection and editing without exposing XPath, raw XML, relationship IDs, part paths, or arbitrary attributes.
- Persist applied plans and mutation footprints with durable OfficeKit tasks so resume reopens the latest reviewed revision and rebuilds its node index instead of restoring a JavaScript heap.
- Establish an immutable three-sample benchmark, deterministic edit matrix, clean-install checks, native rendering, Windows PowerPoint validation, and three independent Agent runs as completion gates.

## Capabilities

### New Capabilities

- `pptx-lossless-edit-plan`: Source-bound inspection, finite Edit Plan compilation, token-preserving execution, mutation-footprint evidence, and fail-closed preconditions for imported PPTX.
- `pptx-lossless-benchmark`: Reproducible real-asset manifests, package and visual oracles, deterministic repetitions, native-host validation, and black-box Agent acceptance.
- `officekit-task-edit-plan-journal`: Durable task storage for compiler operations and reviewed revisions without serializing live JavaScript objects.

### Modified Capabilities

- None. The repository has no root OpenSpec capability catalog; this change adds explicit contracts without replacing the existing presentation object model or wire version.

## Impact

- Extends the existing Office wire v2 schema with an additive internal codec operation and generated bindings.
- Adds a C# PresentationML token editor, JavaScript compiler integration, source snapshots, tests, runtime WASM output, task operation records, Skill guidance, and evaluation evidence.
- Does not add a second PPTX writer, a raw OOXML public API, a universal four-format AST, or a new user-authored file extension.
- Real benchmark assets remain external inputs; the repository records hashes and evidence, not private or third-party source files.
