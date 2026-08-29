## Context

OfficeKit currently exposes presentation creation through a JavaScript object model, MJS/Compose helpers, Help metadata, task authoring plans, and many source-bound edit capabilities. NativeAOT C# already owns the real PPTX package writer and bounded source-preserving edits. The split makes the Agent learn API call order, JavaScript object lifetime, design state, and native editing constraints at the same time, while a fresh context cannot reconstruct the authored program from the PPTX alone.

Earlier changes deliberately avoided a presentation DSL. `presentation-authoring-compiler` retained MJS/Compose as the executable source, `skill-system-reform` excluded a primitive DSL, and `presentation-template-skills` prohibited reference programs and decks. This change supersedes those decisions for OfficeKit 2.0 after the accumulated runtime capability and import evidence made a bounded source language practical.

The shared main worktree contains unrelated user WIP. Implementation uses an isolated worktree from `origin/main` and never stages or resets those paths.

## Goals / Non-Goals

**Goals:**

- Make one strict `.ppj` JSON file the public, durable source for an authored presentation.
- Let an Agent edit PPJ directly and receive precise validation before native bytes change.
- Parse and compile PPJ inside NativeAOT C#, with Node acting only as CLI and framed-transport glue.
- Project every visible third-party object as a typed element or an opaque capability descriptor while retaining unknown OOXML in the source package.
- Preserve third-party no-op bytes and lower supported PPJ changes into bounded source edits.
- Restore an authored PPJ, stable node identity, assets, and design intent from an OfficeKit-authored PPTX.
- Give a fresh Agent a compact PPJ-first Skill and a complete generated language reference.
- Prove the design with one original high-density template and three real workflows before removing the old public authoring surface.

**Non-Goals:**

- Model every OOXML feature as typed PPJ state.
- Expose raw XML, XPath, package part paths, relationship IDs, or arbitrary native fields.
- Put functions, recursion, unbounded loops, network access, or general expressions in PPJ.
- Replace DOCX, XLSX, PDF, or PowerPoint Live models.
- Copy Kimi source text, private implementation, templates, binaries, or proprietary assets.
- Embed projected PPJ into arbitrary third-party packages at the cost of package fidelity.
- Retain a public PPJ/MJS dual authoring system after the 2.0 acceptance boundary.

## Decisions

### 1. PPJ is strict JSON with typed state

The schema identifier is `office-kit/ppj/v1`. A program is one UTF-8 JSON file with root `meta`, `intent`, `design`, `assets`, optional `source`, `components`, ordered `pages`, and presentation-level state such as sections and custom shows.

Elements use a discriminated `type` with shared identity/frame/accessibility fields and type-specific payloads. Simple text accepts a string; mixed formatting uses paragraphs and runs. Ordered page and element arrays are the page order and native stacking order. Unknown fields are rejected.

The checked-in JSON Schema is the public language contract. Generated `ppj.md`, C# discriminated models, examples, and capability ownership are checked against it. JSON object key order is not semantic; array order is.

Alternative considered: a generic `{type, props}` node tree. It was rejected because it hides type errors, weakens generated Help, and recreates the current discoverability problem.

### 2. Template constructs remain finite

PPJ supports named components with typed parameters, slots, variants, finite array repetition, simple equality/presence conditions, local coordinates, and explicit frames. It prohibits recursion, `while`, arbitrary evaluation, dynamic imports, and external effects.

Limits are 16 MiB source JSON, 512 pages, 100,000 expanded elements, 1,024 values in one repeat, and component depth 16. Expansion derives stable instance IDs from component, instance, and item keys and fails before PPTX mutation on collisions, cycles, or budget exhaustion.

Alternative considered: general constraint solving. It was rejected for v1 because fixed presentation canvases can express reusable local layouts without adding a second unpredictable compiler.

### 3. C# directly parses and compiles PPJ

Wire v2 receives additive `PROJECT_PPTX_TO_PPJ` and `COMPILE_PPJ_TO_PPTX` operations. Node reads bounded files, resolves the platform NativeAOT host, and frames requests. C# parses PPJ bytes with `System.Text.Json`, validates typed semantics and assets, expands components, and calls the existing package writer or source-bound edit infrastructure.

No JavaScript Presentation object model is materialized in the PPJ path. The old writer remains reusable internally until the C# PPJ compiler reaches capability parity; it is not part of the final public interface.

### 4. Authored and imported authority are asymmetric

For a source-free program, PPJ is authoritative and PPTX is compiled output. The PPTX contains `/officeKit/program.ppj` plus `/officeKit/program-map.json`, connected by reserved relationships and content types. The map binds program IDs and asset hashes to native object and part fingerprints.

If the embedded program is present, it remains authoritative even after an external editor changes native content. Import restores the embedded PPJ without prompting or merging native drift. A later build writes a new output path and never overwrites the externally edited input. If an editor removes the program parts, OfficeKit performs ordinary projection.

For a third-party PPTX, the original package is authoritative. Import copies it read-only to `<deck>.assets/source/<sha256>.pptx`; `source.uri` is relative and hash-bound. Projected PPJ stays outside the package. No-op build returns the source bytes exactly.

### 5. Projection is complete in visibility, finite in editability

C# projects every visible top-level and nested object into one of two forms:

- a typed PPJ element for semantics the compiler understands; or
- an opaque element with stable ID, page, frame, summary, source revision, and issued `nativeRef` capabilities.

Unknown OOXML does not enter PPJ. At build time C# reprojects the source, compares the baseline projection with the edited PPJ, and lowers supported differences into the existing typed Edit Plan. It re-proves source, target, topology, and expected hashes. Unsupported or ambiguous changes fail closed.

### 6. Assets are local, relative, and content-bound

PPJ references local assets by relative URI, MIME, SHA-256, rights, and accessibility metadata. Build rejects absolute paths, directory escape, URL fetches, missing assets, MIME/magic mismatches, or stale hashes. Image sourcing and generation must materialize files before PPJ compilation.

Authored embedded-program recovery maps PPTX media parts back to declared asset hashes. Third-party source-owned media may remain behind a `nativeRef` until reused or replaced.

### 7. CLI is standalone; Task is optional

The public command family is `officekit ppj import|inspect|check|build|render|review`. The commands work without a Task. `inspect` supports fuzzy discovery and multiple results but never mutates. The Agent edits the PPJ file directly.

`check` validates and optionally performs only deterministic formatting/default repairs. `build` runs structural checking but not render or review. `render` and `review` are separate so the Agent can pay for them at the appropriate stage.

When `--task` is supplied, a successful check or build stores an immutable PPJ revision and descriptor. Resume reopens the latest valid/reviewed revision. There is no duplicate command log; revision diffs provide the audit trail. Legacy `ctx.plan` presentation tasks remain on disk but cannot resume in 2.0.

### 8. The Skill is a short router over one language reference

The main Presentations Skill contains only route selection, source safety, the PPJ workflow, and delivery invariants. `ppj.md` is generated from schema and capability metadata, while focused references cover fonts, shapes, text, charts/tables, media/layers, motion, components/templates, imported native references, scenarios, and review/delivery.

A new Presentation Skill Maintainer and capability registry assigns every existing public Presentation capability to PPJ state, `nativeRef`, compiler/helper, inspect/review, or host-only. A gate rejects orphan capabilities or stale generated language documentation.

### 9. Template schema v3 gains optional executable evidence

Presentation Template v3 keeps `SKILL.md`, preview, and representative examples as its required style contract. It adds optional declared and hashed `referenceProgram` and `referencePptx` entries. The Creator must build both during calibration, but publishes them only when rights and package policy allow it.

This change adds one new `Evidence Ledger` template with original content and geometry. It does not modify Cranberry Evidence or bulk-migrate the template catalog.

### 10. The old public authoring route is removed only at parity

The capability registry is completed before removal. The PPJ path must pass the authored deck, embedded recovery, and complex imported-edit scenarios. Then the package stops exporting public Presentation/MJS/Compose authoring entrypoints and the default Skill removes their routes. External scripts may still generate ordinary JSON.

## Risks / Trade-offs

- [The PPJ schema becomes another incomplete model] → Keep unknown imported objects visible and source-owned through opaque/nativeRef projection; measure user intents rather than OOXML tag coverage.
- [Direct file editing produces invalid intermediate JSON] → Make `check` cheap, precise, and non-destructive; Task snapshots only successful revisions.
- [C# and JSON Schema drift] → Generate the Agent reference and run a registry/schema/model consistency gate.
- [Embedded authored PPJ conflicts with external edits] → Follow the selected PPJ-authoritative policy, preserve the input file, and produce a new output.
- [Component features reintroduce a programming language] → Permit only bounded data iteration and simple conditions; prohibit recursion and arbitrary evaluation.
- [Removing JavaScript breaks hidden callers] → Inventory public exports and package references, reach PPJ parity first, then remove them as an explicit 2.0 break.
- [Template references look derivative] → Use clean-room abstract observations, unrelated content, original geometry/assets, and rights metadata.
- [The migration grows into a test program] → Add only stable format, persistence, source-fidelity, and recovery contracts; use three real workflows for integration evidence.

## Migration Plan

1. Publish the bilingual language rationale and OpenSpec artifacts.
2. Add PPJ schema/model/compiler operations without changing current defaults.
3. Add CLI, projection, source-bound lowering, embedded recovery, and optional Task revisions.
4. Rebuild the Skill, Help, capability registry, Creator, and Evidence Ledger on PPJ.
5. Run the three acceptance workflows in a packed install.
6. Remove old public authoring exports and legacy Skill routes in the same 2.0 branch.
7. Update version, coverage, docs, manifests, package inventory, and release evidence.

Rollback before step 6 is ordinary commit reversion. After the 2.0 break, rollback means restoring the complete 1.x package rather than keeping runtime shims in the 2.0 tree.

## Open Questions

None. Product and compatibility choices are fixed by this change.
