## Why

OfficeKit's Presentation runtime has accumulated a rich native capability set, but the public authoring surface remains split across JavaScript objects, MJS/Compose helpers, task plans, source-bound edit APIs, and long Skill references. A strict declarative program can give an Agent one complete, inspectable, resumable source while allowing the NativeAOT codec to preserve unknown third-party PPTX structures instead of reconstructing them.

## What Changes

- Add PPJ (`PPT Program JSON`), a strict, bounded, single-file JSON language for authored and projected presentations.
- Add a NativeAOT C# parser, validator, component expander, semantic projector, differ, source-bound Edit Plan lowerer, compiler, and embedded-program mapper.
- Add standalone `officekit ppj import|inspect|check|build|render|review` commands with optional Task revision binding.
- Project all visible imported objects as typed elements or capability-bound opaque descriptors while keeping unknown OOXML in the original source package.
- Embed PPJ and stable node mappings in OfficeKit-authored PPTX files; keep projected PPJ outside third-party packages to preserve package fidelity.
- Rebuild the Presentations Skill around PPJ and add a capability registry plus a Presentation Skill Maintainer.
- Extend the Presentation Template schema with optional reference PPJ/PPTX evidence and add one clean-room Evidence Ledger template.
- **BREAKING** Remove the public Presentation/MJS/Compose authoring entrypoints after PPJ capability parity and real-scenario acceptance.
- **BREAKING** Stop resuming legacy `ctx.plan` presentation tasks; they remain on disk and report an explicit unsupported-schema result.
- Publish the completed change as OfficeKit `2.0.0`.

## Capabilities

### New Capabilities

- `presentation-program-language`: PPJ schema, typed elements, bounded components, assets, authoring intent, validation, canonical revisions, and authored PPTX compilation.
- `presentation-program-import`: PPTX semantic projection, opaque/nativeRef representation, source binding, lossless no-op, local Edit Plan lowering, and authored-program recovery.
- `presentation-program-cli`: Standalone Agent-facing import, inspect, check, build, render, and review commands with optional Task integration.
- `presentation-program-skill`: Progressive PPJ guidance, capability ownership, generated language reference, and maintenance rules for future primitives.
- `presentation-program-templates`: Optional reference PPJ/PPTX template evidence, Creator integration, and the Evidence Ledger calibration template.

### Modified Capabilities

None. The repository has no promoted base specs; this change supersedes conflicting decisions in active historical change documents and introduces the replacement requirements as new capabilities.

## Impact

- Public JavaScript exports and Presentation Skill routes change incompatibly in `2.0.0`.
- `proto/`, the NativeAOT Office codec, Node CLI, task store, Help metadata, generated API documentation, template validation, release packaging, and presentation tests are affected.
- DOCX, XLSX, PDF, Office Live adapters, and third-party source bytes remain outside the PPJ authoring migration.
- Existing dirty README, output, and template work in the shared main worktree is explicitly excluded; implementation occurs in an isolated worktree based on `origin/main`.
