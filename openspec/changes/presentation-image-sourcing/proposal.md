## Why

OfficeKit can already place and edit images in PPTX files, but an Agent still has to discover lawful assets, download them safely, preserve provenance, and verify that the delivered deck uses the registered bytes. This missing acquisition layer leaves the default authoring path visually weak and makes rights review ad hoc.

## What Changes

- Add a lazy `officekit image` CLI for candidate search, task-bound acquisition, listing, and PPTX media audit.
- Add Openverse and Wikimedia photo/illustration discovery plus offline Lucide icon discovery, with explicit provider errors and no silent fallback.
- Add task-local, content-addressed image assets, immutable source receipts, search evidence, and attribution metadata without changing the task schema.
- Add bounded HTTPS downloading with destination validation, redirect revalidation, byte/pixel limits, magic/MIME checks, and a strict license allowlist.
- Allow Presentation image placement and Compose image nodes to accept a `FileBlob` directly.
- Add a host-neutral Presentation Skill route for deciding the image role, selecting a compliant candidate, embedding it, rendering it, and auditing crop, clarity, accessibility, rights, and attribution.
- Integrate the additive capability into `2.0.0` without changing the Office wire protocol.

## Capabilities

### New Capabilities

- `image-sourcing-cli`: Stable lazy CLI contracts for search, acquisition, listing, and PPTX image audit.
- `task-image-assets`: Content-addressed task image storage, provenance receipts, secure remote acquisition, and resumable evidence.
- `presentation-image-workflow`: FileBlob image placement plus the Presentation Skill sourcing, review, attribution, and delivery workflow.

### Modified Capabilities

None.

## Impact

- Public CLI: new `officekit image` command group.
- Public JavaScript behavior: existing Presentation image inputs additionally accept `blob: FileBlob`; no new export path.
- Runtime: new leaf-oriented `src/images/` modules that remain unloaded by root import, initialization, template search, and unrelated Office work.
- Dependencies: exact runtime pin for `@iconify-json/lucide@1.2.126`; Openverse and Wikimedia use local leaf-loaded HTTP adapters.
- Skills and release evidence: Presentations reference/workflow, package contents, third-party notices, coverage, and release metadata.
- Unchanged: Office wire version, C# codec, PDF providers, other Office Skills, Live adapters, and template asset format.
