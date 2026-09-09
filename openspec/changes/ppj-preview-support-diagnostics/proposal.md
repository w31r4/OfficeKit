## Why

The preview audit's G-11 remains open: declarations call whole element families supported while visual fields are dropped, chart branches contradict their declared grades, and any diagnostic makes the aggregate partial even when that diagnostic says supported. The local preview needs truthful, field-addressable evidence before its images can safely guide further renderer development.

## What Changes

- Introduce one registry-owned preview support contract for actual schema element/chart types and field behaviors. Derive the existing capability summary from it instead of maintaining contradictory independent lists.
- Evaluate actual input fields, nested elements, page state and inherited/global state conservatively. Unknown or unverified visual state receives an explicit limitation, not an implicit supported grade.
- Add page/element identity, exact input paths, bounded value summaries, stable reason codes and suggested action to diagnostics; retain source-bound/opaque evidence without mutating input.
- Aggregate supported/partial/opaque/unavailable consistently across elements, pages and the complete result, independently of diagnostic count and output publication status.
- Add a separate machine-readable reliability gate: known factual distortions fail it and remain visibly identified in SVG/PNG; approximation and unverified rendering require review. Neither output success nor appearance can clear a failed gate.
- Preserve G-12 publication protection, retained artifacts, source hashes and lazy dependencies. Extend tests and documentation so corrected grades are exercised through actual render results, not only name membership.
- **BREAKING**: previously overclaimed supported grades will become partial/opaque/unavailable where warranted; callers must not treat the exported supported-type set as an inventory of every renderable type.

This change fixes diagnostic correctness, not G-02–G-10 drawing fidelity. It does not mark those gaps or the full renderer goal complete. G-13 page selection, command formatting and automatic review orchestration remain separate; the reliability gate is carried by both in-memory results and published receipts without redefining publication `ok`.

## Capabilities

### New Capabilities

- `ppj-preview-support-diagnostics`: input-sensitive support assessment, traceable limitations, consistent aggregation, factual reliability gating and registry-backed coverage checks for local preview.

### Modified Capabilities

None. The existing G-12 change's output-publication contract remains intact; new diagnostic evidence is additive to it. No main spec currently defines preview support assessment.

## Impact

Targets `src/ppj/svg-preview.mjs`, a small leaf assessment module, `src/ppj/capability-registry.json`, the derived `svg-preview-capabilities.json`, `preview-output.mjs`, capability generation/checking, preview tests and affected documentation. No Office wire change, alternate authoring engine, external service or additional rendering dependency is required. Concurrent errorBars diagnostics must be preserved and incorporated, not overwritten. Implementation must add narrow regression gates and update only verified G-11/G-14/G-15 progress in the audit.
