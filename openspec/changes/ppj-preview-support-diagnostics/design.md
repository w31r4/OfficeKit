## Context

See proposal.md for motivation. Read-only inspection found:

- `svg-preview.mjs` draws directly from canonical program JSON, not the compiler's expanded scene. No diagnostic layer may pretend to have solved G-01.
- `svg-preview-capabilities.json` lists `placeholder` twice, omits actual types such as `ole`, and mixes field/variant names with element/chart types. Chart branches also assign independent statuses.
- `test/ppj-preview-capability-coverage.mjs` merges all grades into one set and scans only top-level fixture elements. `test/ppj-svg-preview.mjs` asserts historical supported membership instead of actual visual state.
- `preview-output.mjs` now separates publication completion from rendering evidence. Its fail helper marks production errors unavailable; its manifest and source preservation must remain intact.
- `scripts/generate-presentation-capability-matrix.mjs` copies a separate preview summary. The authoritative capability registry has no input-sensitive preview contract yet.
- Another workflow is editing errorBars, the capability registry and the preview's fallback chart diagnostic. Merge that diagnostic into the shared assessment rather than reverting it.

## Goals / Non-Goals

**Goals:** a small, local, deterministic assessment boundary shared by in-memory and file output; precise reasons for actual unsupported input state; a reliability gate that cannot be cleared by successful publication or attractive output; one maintained registry source for capability declarations.

**Non-Goals:** no second PPJ semantic compiler, no new scene/layout engine in this change, no blanket conversion of source-bound content into authored primitives, no remote rendering service, no new graphics dependency. The remaining drawing gaps stay open. CLI selection/formatting and orchestration remain G-13; output `ok` continues to mean publication, not overall review acceptance.

## Decisions

### 1. Registry-owned support rules, not new independent type lists

Add an explicit preview section to `capability-registry.json`. Rules record actual schema path/type, semantic owner, assessment handler or accepted state, support grade, reason and regression reference. Keep declarations for real element types and chart types separate from fields (`nativeRef`) and variants (symbol, stream stacking).

Generate the existing `svg-preview-capabilities.json` compatibility summary and include this same information in the presentation matrix. Summary categories are disjoint; a family that is not fully rendered remains partial even if some narrowly defined states of it can be rendered correctly. The runtime evaluates an actual element's fields, not just the family's summary grade.

Add a read-only check mode or equivalent in-memory comparison for generated preview summaries; do not overwrite tracked output merely to test it. Registry rules must resolve to actual schema entries and existing fixtures/tests. Keep rules small and declarative; handler identifiers point to code where conditional interpretation is required.

Alternative rejected: simply mark every type partial in a handwritten array. It fixes some labels but cannot locate omitted fields or prevent the next field from silently becoming supported.

### 2. Assess the real input tree with conservative unknown handling

Use one leaf module, tentatively `src/ppj/preview-diagnostics.mjs`, after normal workspace load/compile and before publication. Walk canonical program/page/element trees with exact JSON paths and IDs, including nested groups and component arguments/definitions where applicable. Treat arrays, empty containers, null, false and zero as actual states, not truthiness defaults.

The mapping distinguishes proven rendered state, explicit non-visual metadata, known limitations, source-owned content and unclassified state. A broad recognized parent must not swallow unknown descendant visual fields. An unclassified field is partial with `preview.field.unassessed`; it cannot be promoted by a successful compiler return. Metadata exclusions must be explicit and tested. Do not dump binary/source payloads or all asset contents into diagnostics.

Inherited state needs explicit handling: theme/styleRef/master/layout/grammar/component expansion that the renderer has not resolved emits limitations at the originating paths and constrains affected nodes/pages. Treat source-bound ownership separately from actual drawing support. Existing source previews remain previews of source state unless edit-to-preview validity is proven.

Known limits include text runs/defaultText, shape geometry and text coexistence, transforms/hidden state, theme and fill/effects, image crop/mask, table sizes/spans, connector endpoints, chart data channels/axes/labels/styles, components, source previews and non-static behavior. Defaults require proof too: absence of a property is not permission to claim unsupported default semantics were honored.

Alternative rejected: independently expand components or resolve all styles here. That duplicates compiler behavior and belongs to G-01/G-04. This module explains the limitation instead.

### 3. Structured diagnostics and deterministic aggregation

Every limitation includes `pageId` and element `id` when applicable, an exact `path`, `status`, stable `reason`, `severity`, `valueSummary`, and `action`. Keep current `id/status/reason` compatibility fields. Value summaries are bounded (for example, 256 characters with truncation indicated), escaped for display and omit source/binary payload contents. Program/page diagnostics have their own exact path rather than invented element IDs.

Assessments expose element/page status plus aggregate status. The deterministic conservative order is `unavailable > opaque > partial > supported`; individual diagnostics preserve why a parent received its grade. Supported informational records do not downgrade an otherwise supported node. An empty list of diagnostics is not proof of support unless the state was actually classified. Remove independent status decisions from drawing branches or normalize them through the same assessment authority.

For bounds, do not announce a complete geometric review. Existing axis-aligned top-level bounds checks remain useful, but nested transforms, text overflow, shadows or masks without resolved bounds receive `preview.bounds.unassessed` with the responsible path. Actual transformed-bounds computation remains in the relevant drawing/layout gap.

Alternative rejected: numeric warning counts as the status. A single missing resource is more consequential than several supported informational records.

### 4. Separate reliability from rendering support and file publication

Expose `reliability: { status, violations }` with `status` equal to `failed`, `requires-review` or `passed`. `failed` means an error-severity factual distortion or unavailable required input/output was detected; `requires-review` means partial/opaque/unclassified visual state remains; `passed` only means these automated diagnostic checks found no limitation in the explicitly assessed state. None of them constitutes human review or editing-fidelity acceptance.

Mandatory factual-error cases are the observed loss of shape geometry when text is present, incorrect connector topology, missing-value coercion or inappropriate bridging, wrong pie proportions, lost hierarchy/flow/OHLC/X/Y/size channels, incorrect cumulative totals, and declared axis/visibility/transform semantics that the drawing contradicts. Register conditional rules and regression cases; distinguish known contradiction from a merely unassessed field. Existing errorBars partial diagnostics remain explicit and at least require review, without pretending error bars are drawn.

Preserve the partially drawn page for debugging, but visibly label pages with a failed gate as unreliable preview and link the warning to diagnostic identifiers. Apply the same warning to the SVG that is rasterized, without changing page dimensions, original element IDs or input state. Explain that a warning does not repair the drawing. Partial/opaque pages also carry a visible review-needed indication so standalone images do not imply complete support.

The publisher copies page/global assessment and reliability into its receipt. A failed factual gate may coexist with `output.status: complete` and publication `ok: true`: files were produced, but are not acceptable correctness evidence. Actual production failures retain G-12's unavailable/incomplete error behavior and also fail reliability. The CLI's existing exit behavior remains publication-only in this change; documentation must say consumers must inspect reliability, not only exit code or `ok`. Integrating overall delivery rejection is part of G-13, not silently implied here.

Alternative rejected: silently suppressing all erroneous charts or calling them opaque. These are authored semantics that still need implementation; a failed gate plus preserved diagnostic output exposes the problem without disguising the missing capability.

### 5. Regression strategy proves assessment behavior

- Test actual nested field paths, page/global inheritance, explicit zero/null/default state, metadata exclusions, unclassified descendants, same element IDs on different pages, unknown types and source-owned payloads.
- Test every aggregation combination, supported informational diagnostics, node/page/document equality, and publication failure dominance.
- Add one proven supported primitive case with geometry/style assertions and a one-field mutation that becomes partial/failed. Do not manufacture support merely to make the test pass.
- Cover each listed factual-error family with a direct assessment/drawing regression. Assert failed reliability and visible warning; these are tests of honest diagnostics, not tests that the chart now draws correctly.
- Run real codec plus sharp for canonical authored and simple source-bound inputs. Confirm reliability is identical in the returned result and persisted receipt, images retain diagnostic labeling, and G-12 hashes/source preservation remain valid.
- Test derived metadata against actual schema enums, disjoint grades and representative nested fields. Add the dependency-light diagnostic suite to appropriate regular gates with segment range tests.

## Risks / Trade-offs

- Conservative assessment initially produces more partial results → document genuine supported subsets and leave the full drawing backlog open; do not lower standards to recover historical green labels.
- Diagnostic volume can grow on large trees → bounded summaries, deterministic deduplication by page/path/reason, no binary traversal; never drop errors silently. Benchmark broader scaling in G-16.
- A new registry property is not equivalent to exhaustive maintenance coverage → validate relevant schema mappings and generated metadata now; broader export/Skill/visual coverage enforcement stays G-15.
- Source-bound/global styles can be difficult to attribute → report the actual owning path and affected page instead of claiming fully resolved semantics.
- Concurrent modifications touch shared files → inspect the latest diff immediately before patching, preserve new field behavior and rerun the affected tests.
- Output callers may still read only `ok` → explicitly document its publication meaning and the new required reliability check; do not claim G-13 workflow integration is done.

## Migration Plan

Land registry rules and tests alongside runtime assessment; regenerate compatibility summaries with the repository script. Replace old tests that equate type membership with full support with actual-state assertions. Update output docs and the audit with commands and evidence, keeping all drawing gaps open. No wire/package upgrade or destructive source migration is involved. Do not revert G-12 publication protection if a newly accurate diagnostic breaks a historical assertion.
