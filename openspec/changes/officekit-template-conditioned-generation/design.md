## Context

An imported PPTX is both a visual reference and a source-bound OPC graph. A
profile can describe its design language, but it must not grant permission to
rewrite unknown XML. New pages therefore start as proven source-slide clones;
content changes stay within typed, inherited leaves and the original graph is
retained as the source of truth.

## Decisions

### 1. Separate exact following from conditioned generation

Exact following uses a complete frame map and edits declared inherited targets.
Conditioned generation selects a small set of source archetypes for a new
narrative, allowing a source slide to be reused. It does not select a generic
preset after a user supplied a stronger template.

### 2. Profile is evidence, not authority

`designProfile()` reports the source revision, canvas, palette, typography,
density, archetypes, components, and opaque summaries. Candidate status and
clone capability decide whether a source slide may be copied. Blocked or
inspect-only components remain available for inspection but are never flattened
to make generation succeed.

### 3. Clone across boundaries and locate by source ordinal

The codec allows one pending clone per source origin. Each round crosses an
export/reimport boundary. Public slide ids are regenerated from display order,
so the generation runner carries the frozen source ordinal and clone occurrence
through the transaction and resolves the final manifest from those locators.

### 4. Preserve inherited semantics

Text replacement is run-scoped and must keep font, paragraph/run topology,
geometry, placeholders, relationships, and opaque descendants. If copy does not
fit, the route chooses a shorter value or another archetype. It never silently
shrinks typography or reconstructs a slide.

### 5. Review against the source baseline

The generated artifact is exported and imported a second time. Target values,
package parts, source hash, `verify`, and `validateLayout` are recorded. Existing
source issue categories are reported as inherited; a new issue category is a
regression. A renderer refusal is recorded as `visualReview: "unavailable"`
rather than replaced by a fake preview.

## Evidence Matrix

The frozen external samples are 算秩未来 (21 slides), 蓝灰酸性模板 (19 slides),
and 麦肯锡风客户忠诚度 (8 slides). The deterministic benchmark generates 10
new slides per sample. The independent black-box lane runs a fresh three-phase
Agent task per sample (profile/plan, one bounded clone edit, and separate
review), then checks the packed install, second import, source protection, and
non-target package parts. Its compact evidence lives in
`evals/pptx-generation/agent-blackbox.v1.json`; visual review is reported as
unavailable when the portable renderer cannot consume a source's geometry.
Native Windows PowerPoint remains a separately scoped host gate.
