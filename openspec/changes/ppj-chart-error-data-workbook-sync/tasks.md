## 1. Formula binding and synchronized values

- [x] 1.1 Add formula projection and immutable binding checks; verify no-op, source-free rejection and formula topology rejection in focused PPJ tests.
- [x] 1.2 Implement owned workbook planning, worksheet token splicing and outer package footprint/hash propagation; verify ordinary, combo and grouped-chart updates, repeat projection and unchanged unrelated bytes.
- [x] 1.3 Reject shared/missing/overlapping/stale/dependent workbook cases and verify no file is emitted; rerun literal/scalar error-bar regressions.

## 2. Documentation and evidence

- [x] 2.1 Update the field reference, registry and F-07; regenerate/check derived docs, run strict OpenSpec and focused checks, and record exact results and remaining topology/visual limits.

Evidence (2026-09-09): focused managed error-workbook, literal/scalar lifecycle,
shared error-bar codec and existing chart-cache/workbook leaf regression passed
32/32. The ownership rejection additionally exercises direct artifact export so
the underlying native path cannot edit an unrelated workbook-backed channel.
The new grouped-chart fixture verifies ChartPart/workbook-only changes without
rewriting its slide. Tests use the original source bytes and fresh projections.

Presentation Skill maintenance check passed (151 Help APIs, 278 native leaves),
the generated PPJ manual and capability matrix were regenerated, reference Skill
sync passed (333 files), and strict OpenSpec validation and diff checks passed.
Generated release docs were produced in an isolated checkout to keep the
concurrent preview diagnostics work out of this atomic field change.

Broader checks remain limited: skill-portability still fails its existing
powerpoint-live-control REPL expectation. The additional
GroupShapesAuthorImportNestedModeledChildrenAndValidateNativeTree run reaches
the unrelated irregular-group Source.Editable assertion at PptxCodecTests.cs:6209
and fails there; the group author/edit/round-trip checks before it pass. The
opaque placement-editability path is unchanged by this change. This is managed
structural/round-trip evidence, not a NativeAOT release or PowerPoint/visual test.
Formula topology changes and the full F-07 backlog remain open.
