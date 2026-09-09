## Context

See proposal.md for motivation. Read-only investigation on 2026-09-09 found these concrete boundaries:

- `PpjAuthoredPresentationCompiler.Compile` creates `AuthoredSourceFreeBuildPlan` from `validation.Expansion` and `Catalog`, then returns `validation.CanonicalJson` in `ProgramJson`.
- `AuthoredSourceFreeBuildPlan.MaterializeSlide` consumes expanded element models/JSON, invokes `BuildElement`, applies layout text precedence, and produces the slide the writer uses. `PptxCodec` invokes it once per slide and releases ordinary previous-slide graphs; Morph selectively retains the previous slide.
- `Catalog` resolves theme/grammar colors, fonts, named styles, layout/master text precedence and native asset IDs. The native asset identity is derived from MIME/hash, not necessarily the public PPJ asset ID.
- `ChartCompiler.BuildInto` already lowers heatmap, candlestick, treemap, sunburst, Sankey, stream, pictograph and selected numeric charts into `PresentationGroup` primitives. Other charts remain native `PresentationChart` state. Reusing these results is materially different from merely expanding components.
- `PpjExpandedNodeModel` already records page, ID, source ID, component/instance/repeat key, actual program path and z-order. Writer-created children need attribution to their semantic owner in addition to this map.
- Source-bound compilation has physical no-op, precise native-leaf edit-plan and semantic export branches. Precise edit-plan output is not guaranteed to be represented by the earlier in-memory `presentation`; taking that object alone could show old state.
- `PresentationProgramRequest/Result` currently have no scene option or result. Both `PpjCodecProtocol` and `CodecProtocol` handle compile responses. JavaScript `native.mjs` selectively copies PPJ receipt fields and `workspace.mjs` currently only forwards `includeNodeMap`.
- G-11/G-12 now provide structured diagnostics, independent reliability, visible warnings, source/input/candidate hashes and non-overwriting publication. `test/ppj-svg-preview.mjs` exercises real codec/sharp authored/source-bound output. Native lifecycle tests demonstrate source re-projection and unchanged-part assertions.

## Goals / Non-Goals

**Goals:** one compiler-owned, read-only scene boundary used by local preview; equivalent high-level and explicit authoring share effective geometry/data/style; candidate-bound source previews; exact provenance and conservative residual limitations; no extra scene cost in ordinary compilation.

**Non-Goals:** no new editable language, public JavaScript Presentation facade, second OOXML parser, standalone layout solver, network service or rendering dependency. This is not a promise that all native presets, text layout, chart ticks, effects or host behavior are already painted. Those remain G-02–G-10, with exact limitations rather than field loss. G-13 CLI routing and G-16 global performance targets remain separate; preserving ordinary compile resource behavior is required here.

## Decisions

### 1. Reuse the native writer IR in an opt-in envelope

Add `include_preview_scene` to `PresentationProgramRequest`, default false, and a dedicated typed preview envelope to `PresentationProgramResult`. The envelope contains a `PresentationArtifact`, scene format version, origin (`authored-lowering` or `candidate-import`), canonical program hash, actual candidate hash, semantic/native bindings and native asset identity references. Scene bytes have a deterministic digest recorded in the preview output receipt. Reuse existing protobuf field definitions rather than maintain a parallel visual schema.

The scene is read-only evidence, not an export-authorized `ArtifactEnvelope`: do not attach source package bytes, writable ownership tokens or opaque OPC payloads as an additional transport copy. Original source bytes and the public canonical PPJ remain unchanged and retain their existing meaning. Opaque native descriptors remain identifiable and limited; sanitizing package payload transport must not erase their visible presence or pretend they are editable.

Allocate unused field numbers after inspecting concurrent proto changes. This is an additive option/result on the existing PPJ operations; keep the current Office wire version unless a separately demonstrated compatibility break requires a coordinated migration. A scene format version and presence check are mandatory: an older codec that ignores the request must produce a clear unavailable/rebuild-required preview result, never a silent canonical-PPJ fallback. Reject `validationOnly + includePreviewScene` and projection requests carrying the compile-only option explicitly.

Alternative rejected: change `programJson` to expanded JSON. That changes the canonical revision/hash contract, still misses catalog lowering, and makes editable PPJ ambiguous. Reusing `CodecResponse.artifact` without a read-only envelope was also rejected because it mixes legacy export authority and potentially huge source-package snapshots into the PPJ path.

### 2. Capture authored slides at the writer boundary

Give the authored plan an optional scene collector. Capture the materialized slide that is actually handed to the writer, after component expansion, catalog resolution and chart lowering; preserve slide order, nested group/child frames, hidden state, master/layout context and typed field presence. Do not call expansion or `BuildElement` a second time for preview. Keep existing Morph previous-slide behavior and ordinary scene-disabled lifetime intact.

The scene collector retains one opt-in immutable snapshot, with deterministic page/node ordering. Bound scene size and node traversal using existing codec/expanded-element budgets; budget overflow is an explicit error, never a truncated scene reported complete. The unchanged output PPTX hash with the option on/off is part of acceptance.

Capture complete native visual state, including fields not yet supported by the SVG painter. Compiler-resolved styles are authoritative; residual inherited native state such as theme transformations or host-resolved placeholder details is carried with an explicit unresolved scope, not replaced by white/default styling. Pixel text layout remains a painter responsibility, not falsely declared solved by materializing the slide.

Alternative rejected: re-import every authored PPTX. It loses the direct expansion-to-writer provenance, adds a second package parse to ordinary authored preview, and can obscure the compiler's generated grouping. Authored capture is the direct path.

### 3. Source-bound scenes use the exact candidate

After source-bound compilation succeeds, use the existing native PPTX importer on the resulting bytes for the scene; a physical no-op uses the exact original bytes. Cover the leaf-only edit-plan and semantic export paths, including the direct native-host file-backed `ReuseSourceFile` branch. Do not restore the embedded PPJ snapshot as if it were a freshly read visual model. This import is read-only and opt-in; it must not re-export the candidate or mutate the source.

Bind the scene to `OutputSha256`, not merely `SourceSha256`. Preserve semantic IDs using the existing native bindings/part identities; track imported/generated native IDs separately where needed. If a scene node cannot be safely attributed, keep a real page/owner path and a separate scene path, explicitly mark attribution unavailable, and do not invent an editable PPJ child path. Assets referenced by the candidate scene come from the exact imported candidate and must resolve by native identity/hash.

Keep source/opaque states distinct from supported authored primitives. A view of a source preview image is not flattening the original object, but neither is it proof of updated opaque payload content. A payload-only OLE edit whose source thumbnail is unchanged must remain visibly limited.

Alternative rejected: return the source-bound compiler's existing mutable `presentation`. The native-leaf fast path can modify the final bytes without updating every modeled field, which would violate the actual-candidate requirement.

### 4. A leaf scene adapter feeds the SVG painter

`native.mjs` and `compilePpjWorkspace` forward the explicit option and scene. `renderPpjToSvg` requests it, verifies scene version/program/candidate identity, and draws the scene rather than interpreting component repeat/layout or dataset/encoding in JavaScript. It must not import the retired public Presentation facade.

A small scene adapter owns mechanical representation conversions: EMU/angle/font units, optional-value presence, native asset IDs, typed content cases and property paths. It does not expand components, evaluate grammar tokens, choose data channels, regenerate chart topology or guess missing observations. Do not round-trip the scene through editable PPJ or a lossy pseudo-PPJ schema just to reuse the old drawing branch.

Route native shape/image/group/table/connector/chart and remaining content cases explicitly. Reuse common SVG paint helpers where their semantics agree. Generated vector charts consume their actual native child shapes and paths rather than return to the old chart heuristics. Native ChartPart types consume their native data/axes; unimplemented painting fields remain diagnosed, not omitted by the adapter. This scene contract must preserve every available native field even when the painter is partial.

The first equivalence fixtures must actually paint distinguishable compiler-resolved positions and styling, not merely emit identical opaque boxes. Arbitrary preset/effect completeness is separate, but an implementation that transports a scene and continues drawing canonical guesses does not complete G-01.

### 5. Preserve G-11 provenance and honest support

Assess both the original input/ownership boundary and the scene's actual visual field consumption. Use the original PPJ `path` for source diagnostics and a separate `scenePath` for lowered/generated fields; bind generated children to their real semantic owner. IDs and z-order must remain traceable through component instances and repeated items.

G-11 factual rules currently describe the old canonical drawing implementation. Make rules renderer-profile/state-sensitive: for example, do not retain a `parents ignored` factual error after the native treemap layout and its actual geometry are used. Conversely, merely receiving a scene is not evidence to clear text, axis, geometry or missing-point failures. Each removed limitation requires a specific native-state-to-SVG regression; unrelated opaque, unsupported and factual limitations retain their severity.

Register scene owner/mapping evidence in the existing capability registry. Use native descriptors/current wire state to reject unclassified content and diagnose unknown visual descendants; do not introduce a competing handwritten feature inventory. Extend `render.json` with scene origin/version/digest and candidate linkage, and keep the visible warnings, publication-only `ok`, source hashes and artifact hashes intact.

### 6. End-to-end evidence, not only a transport test

Native tests compare scene-on/off compilation and the captured slide against the writer input. Use paired high-level/explicit fixtures for nested components, repeats/slots, named style/grammar resolution, dataset/encoding and both vector/native charts. Compare geometry, ordered data including null/zero, styles and asset identity while normalizing only explicitly documented generated-ID differences.

JavaScript tests must show the real painter used changed compiler coordinates/colors/data and never called old component heuristics. Cover unknown fields, future scene versions, missing scene, corrupted linkage, mismatched assets, scene budgets and lazy imports. Retain G-11/G-12 failures and source preservation tests, and verify both returned and persisted scene/reliability evidence.

Source-bound regression includes no-op exact bytes, one native-leaf text/frame edit and one semantic edit, checking the scene against a fresh import of the actual candidate. Preserve opaque sibling ordering and non-target ZIP parts; do not count authored round-trip snapshots alone as source-bound proof.

Rebuild the NativeAOT runtime actually used by the JS integration test. Run proto generation/checks and review generated diffs; `proto:check` invokes generation plus a git-diff check, so expected pending binding changes are not equivalent to regeneration drift. Verify generation determinism and report the exact command outcome without staging unrelated changes merely to make that check green.

## Risks / Trade-offs

- Larger preview responses and retained slide graphs → explicit opt-in, bounded scene payload, unchanged scene-disabled lifecycle, representative size/time measurements; do not claim full G-16 performance acceptance.
- Source-bound extra parse → only for requested previews; reuse the existing codec and exact candidate bytes, with measured cost and no second authoring pass.
- Native scene still contains host-resolved state → preserve the state and report its exact limit; do not confuse writer IR with a pixel-complete display list.
- Generated IDs and native asset IDs differ from user IDs → explicit bindings and deterministic provenance tests; never infer edit authority from scene IDs.
- Existing factual rules can become stale after scene adoption → profile-sensitive mappings with direct regression evidence before removing warnings.
- Concurrent proto/registry/compiler edits → re-read shared diffs before each patch, regenerate using current schema and retain the other workflow's tests.

## Migration Plan

Introduce the option/result and native capture/import first with default-off tests; regenerate and rebuild the matching runtime. Then wire the leaf consumer and G-11/G-12 evidence. Switch local preview to require the new scene only when end-to-end tests pass; do not offer a hidden fallback to old guesses. Ordinary PPJ check/build and original canonical JSON remain compatible. If a release must be rolled back, roll back the preview consumer and matching codec together; never rewrite user PPJ or source files for compatibility. Update the audit only for behavior actually demonstrated.
