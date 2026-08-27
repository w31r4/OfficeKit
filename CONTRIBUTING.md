# Contributing to OfficeKit

Thanks for helping improve OfficeKit. The project is English-first in its
source, public API, and release records; Chinese documentation may accompany
the English version where it helps users.

## Before changing code

Read [`AGENTS.md`](AGENTS.md), inspect `git status -sb`, and identify whether
the change belongs to the JavaScript runtime, C# codec, PDF provider, Excel
host adapter, Skill package, or release tooling. Keep unrelated worktree
changes intact. Use an OpenSpec change under `openspec/changes/` for a feature
that changes a public contract or spans multiple layers.

## Repository layout

The public runtime is in `src/`; the C# OfficeKit Codec source is in
`native/OfficeKit/`; wire contracts are in `proto/`; generated runtime assets
are staged into target-specific `packages/office-kit-codec-*/`; and canonical
Agent Skills are in `skills/`.
The root `.claude-plugin/marketplace.json` and each plugin's
`.codex-plugin/plugin.json` are distribution metadata, not alternative Skill
implementations.

## Skill and plugin changes

- Keep one canonical Skill tree under its owning `skills/<plugin>/` directory.
- Keep frontmatter, examples, references, and agent metadata portable across
  Claude Code, Codex, and other hosts.
- Do not require a host-specific tool, thread ID, MCP server, cloud drive, or
  image-generation tool. If visual input or generation is unavailable, use the
  documented structural QA path and report that visual review is unavailable.
- Update the relevant `.codex-plugin` metadata and the Claude marketplace
  entry when a plugin is added, renamed, or its release version changes.
- Preserve the MIT license and byte-integrity records for the default template
  library; do not mix its assets into AGPL source files.

## Runtime and codec changes

Office DOCX/XLSX/PPTX authoring uses the canonical OfficeKit Codec path. Do not
introduce a second JavaScript Office writer or an implicit fallback. PDF
providers remain explicit capability packs with their own licensing and
runtime requirements. Imported content that cannot be represented safely must
be preserved as opaque data or rejected before mutation.

For changes that cross JavaScript, protobuf, C#, and NativeAOT, update the model,
wire contract, native codec, generated runtime, public API/help, Skill, and
tests together. Keep source-bound edits immutable, atomic, and auditable.

## Validation

Run the narrow checks first, then the full suite when practical:

```sh
npm test
npm run docs:api
npm run proto:check
npm run build:office-kit
npm run verify:office-kit-build
node test/claude-plugin.mjs
node test/package-contents.mjs
```

If a provider, renderer, Office host, or platform runtime is unavailable,
report the exact skipped gate and the environment needed to run it. Do not
replace a missing real-provider check with a fake success.

## Commits and pull requests

Keep commits small and coherent. A public feature should normally include its
implementation, tests, Skill/docs updates, and generated release evidence in
one atomic change. Never rewrite shared history or force-push. Commit messages
should state the user-facing change; include the repository's required
`Co-Authored-By: Enter Code <noreply@enter.pro>` trailer when committing through
the project workflow.

Pull requests should explain the changed contract, source/provenance boundary,
tests run, platform/provider skips, and any remaining partial coverage. Do not
claim native host or visual acceptance from a mocked test alone.

## Licensing and provenance

OfficeKit is AGPL-3.0-or-later. Third-party code, reference Skill material,
OfficeKit Codec compatibility work, MuPDF, and the MIT template library
retain their declared licenses and notices. New assets must include their
license and provenance before they enter a published package.
