## 1. Non-destructive artifact publication

- [x] 1.1 Reserve a fresh destination and publish page artifacts exclusively using safe unique filenames; verify existing directory/file/symlink rejection, protected input preservation, concurrent destination acquisition and unsafe/case-colliding page IDs in `test/ppj-preview-output-evidence.mjs`.
- [x] 1.2 Implement pending/final output lifecycle with retained partial results and structured publication errors; verify first/later-page raster failures, artifact write failure, pending/final manifest failure and marker cleanup behavior without overwriting or recursively deleting evidence.

## 2. Truthful receipt and resource snapshot

- [x] 2.1 Add the versioned actual-artifact receipt with SVG/PNG hashes and sizes, input/canonical/candidate/source/asset identity, runtime/backend identity and distinct production state; verify every advertised artifact against its bytes and omit all failed PNG compatibility references, including a source-bound receipt case.
- [x] 2.2 Use loaded workspace asset bytes for preview embedding and receipt hashes; verify a controlled after-load path mutation cannot change compiled-versus-preview assets and that absent/inconsistent asset IDs do not become silent empty hrefs.
- [x] 2.3 Keep raster loading lazy and represent missing dependencies separately from raster execution failure; verify retained SVG-only incomplete evidence with non-successful publication, no automatic fallback/download, and unchanged in-memory SVG behavior without loading the PNG dependency.

## 3. Integration, documentation and audit evidence

- [x] 3.1 Update the existing preview smoke to use a fresh child destination and preserve its successful page filenames; run `node test/ppj-svg-preview.mjs`, `node test/ppj-preview-output-evidence.mjs` and `node test/ppj-preview-capability-coverage.mjs`, with real codec/sharp success in addition to injected failure cases.
- [x] 3.2 Register the dependency-light output evidence regression in package scripts and the appropriate regular test gate; run `node test/gate-policy.mjs` and the new registered test, confirming no root eager renderer dependency was introduced.
- [x] 3.3 Document fresh-directory usage, retained incomplete outputs, truthful manifest fields and error handling; update only verified G-12 audit progress, preserve all other open gaps, and run `git diff --check` plus `openspec validate ppj-preview-output-evidence --strict`. If canonical Skill/reference files change, also run `node test/skill-portability.mjs` and `node test/reference-skill-sync.mjs`.

Verification note (2026-09-09): `node test/gate-policy.mjs`, `npm run test:ppj-preview-output`, `node test/ppj-svg-preview.mjs`, `node test/ppj-preview-capability-coverage.mjs`, and strict OpenSpec validation pass. The gate contract now checks the active PPJ scripts, rejects retired Presentation test entry points, and verifies contiguous segment coverage and each segment's first test. This is focused output-publication validation, not full-suite or host-PowerPoint acceptance.
