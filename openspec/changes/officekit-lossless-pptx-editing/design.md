## Context

The current importer projects PresentationML into an Agent-friendly `Presentation` object while retaining a trusted source-package snapshot and source bindings. Existing typed export remains valuable for source-free authoring and broad modeled edits, but reconstructing a changed SlidePart can normalize namespace declarations, XML lexical form, and unmodeled descendants. It can also inspect unrelated shapes through a smaller authoring grammar before reaching an otherwise safe target.

The first implementation milestone has already established the critical feasibility result on three third-party files: a proven no-op can return the original package bytes, and two complex presentations can change one `a:t` leaf while every non-target OPC part and all masked target XML bytes remain identical.

## Goals / Non-Goals

**Goals:**

- Separate lossless source ownership, semantic projection, and executable mutation intent.
- Make every accepted edit source-bound, finite, deterministic, auditable, and independently re-proved by the C# codec.
- Preserve unknown Office structures without requiring their full semantic modeling.
- Let a fresh Agent resume from reviewed artifact bytes and durable operation evidence.
- Define completion through real packages, renderers, PowerPoint, clean installation, and independent Agent runs.

**Non-Goals:**

- No universal AST across PPTX, DOCX, XLSX, and PDF.
- No public XPath, arbitrary XML/attribute editing, relationship rewriting, or raw OOXML fallback.
- No claim that every PresentationML construct becomes typed or editable.
- No automatic restoration of JavaScript heaps across task sessions.
- No completion claim from unit tests or a single successful sample.

## Decisions

### 1. Three IR layers with different ownership

The lossless source graph owns exact OPC bytes, relationships, raw XML token positions, and unknown native subgraphs. The existing `Presentation` model remains the semantic projection used by Agents. A finite Edit Plan records only authorized mutations and preconditions. This avoids making the semantic model carry lexical XML details and avoids presenting an opaque XML tree as an Agent API.

Alternatives rejected: a complete OOXML JSON AST would move the schema surface rather than remove it; HTML would simplify rendering but abandon source fidelity; full SlidePart reserialization cannot prove masked byte identity.

### 2. No-op is an exact source return

An imported presentation may use the source package directly only after the complete presentation projection, not merely slide elements, is proven unchanged. Pending clones, changed custom shows, sections, comments, notes, view state, or any other presentation-level state force the existing typed path. This makes no-op equality a compiler proof rather than a ZIP normalization promise.

### 3. C# re-proves and token-splices every operation

`APPLY_PPTX_EDIT_PLAN` receives the exact source SHA-256 and bounded operations. Each operation binds slide part, slide XML hash, native shape-tree index, element hash, semantic hash, leaf ordinal, and old text hash/value. Open XML SDK parsing is a read-only structural oracle. Mutation occurs in the original UTF-8 XML token stream and changes only the selected `a:t` token plus a necessary `xml:space` attribute.

The codec rejects stale source, stale element, unsupported target kind, overlapping operations, duplicate targets, invalid paths, or output scope drift. It reopens the result and validates that only planned parts changed and that every new leaf is present.

### 4. Native leaves remain capability-issued

`inspect({ includeNativeLeaves: true })` will issue revision-bound leaf IDs only for codec-proven safe fields. `editNativeLeaf` accepts one such ID, expected hash, and value. It never accepts a part path, XPath, QName, arbitrary attribute, relationship ID, namespace, identity, or topology mutation. Initial leaf kinds are text, color, and local geometry scalars; unsupported nodes remain inspectable only through summaries or opaque preservation.

### 5. Task operation records are immutable evidence

Applied plans are stored under `.office-kit/tasks/<task-id>/operations/` with private permissions and linked from the reviewed commit. The record includes source and output revisions, requested preconditions and values, generated footprint, changed parts, and hashes. Resume restores the latest reviewed artifact bytes and rebuilds its node index; source text in the REPL journal is audit evidence, not executable state replay.

### 6. Benchmarks use independent oracles

The three external PPTX files are identified by hash in a versioned manifest. Oracles compare uncompressed OPC entry bytes, relationship graphs, masked target XML, native structure counts, second import, render pixels for non-target pages, and PowerPoint behavior. Each edit runs three times from a clean source and must produce identical output and footprint hashes. Kimi, HTML, and PPTD results are context, not acceptance standards.

## Risks / Trade-offs

- [Raw token location diverges from Open XML object order] → Bind both native indices and hashes, compare SDK and token targets, then reject ambiguity.
- [Protobuf schema growth changes message wire ordering] → Use typed semantic equality for internal projections while keeping package and XML byte equality as the artifact oracle.
- [A safe operation needs extra dependent parts] → The compiler declares the footprint before execution; the codec rejects any undeclared changed part.
- [ZIP container metadata changes after a real edit] → Require exact no-op package bytes and exact uncompressed non-target OPC contents; report container-level differences separately.
- [Real assets cannot ship in the repository] → Store hashes, inventories, target selectors, and evidence; require local assets for the slow benchmark lane.
- [Renderer equality misses Office-specific repair behavior] → Keep Windows desktop PowerPoint open/browse/save-copy acceptance mandatory before completion.
- [Controlled native editing becomes raw XML by another name] → Capability-issued leaf IDs and a fixed leaf-kind registry prevent arbitrary paths and topology edits.

## Migration Plan

1. Land exact no-op return, text-leaf Edit Plan, C# token executor, generated bindings, regression tests, and first benchmark evidence.
2. Add durable task operation records and controlled native-leaf inspection/editing.
3. Expand the deterministic real edit matrix without weakening preconditions or falling back to whole-part reconstruction.
4. Run clean-install, full repository, reproducible WASM, hosted CI, renderer, Windows PowerPoint, and Agent 3/3 gates.
5. If the new path must be rolled back, remove its compiler dispatch; the existing typed exporter remains independently available and the wire version remains 2.

## Open Questions

- Which color and local geometry leaves have enough lexical and semantic evidence for the second native-leaf tranche?
- Which SmartArt text leaf in the real benchmark has a stable package-local binding without relationship or topology mutation?
- Which Windows runner and PowerPoint version will hold the signed final host-acceptance evidence?
