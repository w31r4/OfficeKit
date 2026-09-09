## Purpose

Provide local presentation previews with compiler-owned visual state and exact candidate provenance, so high-level authoring and source-bound edits are reviewed without a second conflicting interpretation of their meaning.

## ADDED Requirements

### Requirement: Compilation offers an explicit read-only preview scene

PPJ compilation SHALL provide an opt-in versioned scene without changing canonical program JSON, ordinary compilation output, or the meaning of its existing hashes. Scene-disabled build/check calls SHALL NOT retain or transport a preview scene. Validation-only or projection operations SHALL reject a compile-only scene request explicitly.

#### Scenario: Existing callers do not request a scene

- **WHEN** a caller builds or checks a program without the preview option
- **THEN** existing results remain scene-free and preview collection does not change file bytes or ordinary slide-lifetime behavior

#### Scenario: Preview is requested from an incompatible runtime

- **WHEN** the returned scene is absent, unsupported, malformed or does not match the compiled revision/candidate
- **THEN** preview reports an explicit unavailable result rather than silently drawing canonical guesses

### Requirement: Authored preview consumes the same resolved meaning as export

An authored scene SHALL preserve the expanded geometry, ordering, resolved styles, asset identities and chart data or generated primitives consumed by export. Equivalent high-level and explicitly expanded inputs SHALL produce equivalent scene geometry, data and styling, allowing only documented semantic-ID differences. Preview SHALL NOT independently reinterpret components, grammar or data encoding.

#### Scenario: Repeated component uses a named style and token

- **WHEN** a component with repeats, slots and nested content resolves layout and named/grammar styling during compilation
- **THEN** its preview uses those exact resolved positions, ordering and styling and remains attributable to the original component instance and source fields

#### Scenario: Encoded data or a vector-lowered chart is previewed

- **WHEN** compilation resolves a dataset/channel mapping or produces a chart's native vector children
- **THEN** the scene retains that exact data/topology and those children, including null versus zero, instead of generating a different chart model

### Requirement: Source-bound scenes represent the actual output candidate

A source-bound scene SHALL be bound to the actual candidate bytes, including precise native-leaf edits and semantic edits. A no-op SHALL retain exact original file identity. Scene generation SHALL NOT rewrite the source/candidate, flatten opaque objects, restore stale authored snapshots as current visual evidence, or imply editing authority for scene nodes.

#### Scenario: A precise edit changes only candidate native bytes

- **WHEN** an imported text or frame leaf is changed through its issued capability
- **THEN** preview reflects the actual changed candidate state and carries the candidate hash, while original source bytes and non-target content remain preserved

#### Scenario: Opaque siblings or stale source thumbnails remain

- **WHEN** the candidate contains unmodeled native content or a source preview that cannot establish changed payload appearance
- **THEN** its presence and ordering remain visible and limited, with no fabricated supported state or flattened edit result

### Requirement: Scene transport preserves visual state and traceability

Scenes SHALL identify their version, origin, program/candidate identities and deterministic content digest. Each drawn or limited node SHALL retain page and semantic ownership, with separate scene paths for generated content. Scene assets SHALL resolve to the exact compiler/candidate bytes by native identity and hash. Unknown fields, unresolved inheritance and unpainted content SHALL remain explicitly classified; the adapter SHALL NOT silently drop them.

#### Scenario: Native asset IDs and generated child IDs differ from input IDs

- **WHEN** lowering produces native identifiers distinct from user-authored identifiers
- **THEN** the renderer resolves the correct asset bytes and attributes every generated child to a real semantic owner without inventing editable input paths

#### Scenario: A new visual field or content type lacks painter support

- **WHEN** the scene contains visual state without a verified paint mapping
- **THEN** the field remains in scene evidence and causes a traceable limitation rather than a complete-support claim

### Requirement: Local preview actually draws the compiler scene

Local preview SHALL consume the scene through the independent local runtime without introducing another authoring engine or eager specialist dependency. The equivalent-input tests SHALL demonstrate actual rendered geometry and styling, not only equal placeholders or equal serialized receipts. Remaining drawing gaps SHALL stay explicit, and previous factual failures SHALL be cleared only when the affected real mapping is verified.

#### Scenario: Compiler-resolved placement differs from the old heuristic

- **WHEN** a component's compiler layout differs from the old label/value vertical arrangement
- **THEN** the preview shows the compiler-resolved arrangement and the old heuristic is not used

#### Scenario: Scene availability does not repair a painter limitation

- **WHEN** the scene is valid but a text, geometry, chart, effect or missing-observation behavior remains incorrectly painted
- **THEN** the relevant limitation or failed reliability persists with its visible warning; scene receipt success cannot clear it

### Requirement: Publication and resource boundaries remain intact

Scene-based preview SHALL preserve independent support, reliability and file-publication states, non-overwriting output, truthful retained artifacts and matching source/scene/image hashes. Scene collection and consumption SHALL obey explicit node/byte budgets; exceeded budgets SHALL fail without claiming a complete truncated scene. No in-memory SVG call SHALL load a raster backend merely to inspect the scene.

#### Scenario: Scene or publication fails

- **WHEN** a scene budget, integrity check, required asset, raster dependency or output write fails
- **THEN** the failure remains explicit, existing files and inputs are preserved, available partial evidence is truthful, and reliability cannot report passed

#### Scenario: Scene preview publishes a partially supported page

- **WHEN** the requested files publish successfully but unpainted native fields remain
- **THEN** returned and persisted scene/support/reliability evidence agree and standalone images retain the appropriate review warning
