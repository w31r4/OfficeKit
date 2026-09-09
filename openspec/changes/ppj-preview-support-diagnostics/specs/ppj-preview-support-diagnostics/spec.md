## Purpose

Make local PPJ preview support claims traceable to actual input state, distinguish factual rendering failures from visual limitations, and preserve that distinction in both images and machine-readable evidence.

## ADDED Requirements

### Requirement: Support declarations describe the actual language and implementation

Preview capability declarations SHALL distinguish actual element/chart types from variants and source-binding fields. A capability SHALL have one summary grade; declarations SHALL NOT call a whole family supported when its declared visual state is only partially rendered. Registry-owned support evidence and the published summary SHALL agree.

#### Scenario: Conflicting or obsolete declarations

- **WHEN** the summary contains a duplicate grade, an obsolete element name, a variant presented as a schema type, or a state stronger than its recorded mapping supports
- **THEN** its consistency check fails and the renderer does not infer complete support from that declaration

### Requirement: Actual fields are assessed without silent loss

Preview SHALL assess the supplied program, pages and nested elements, including global/inherited state and explicit zero, false, null and default semantics. Every unimplemented or unclassified visual field SHALL produce a limitation. Recognizing a parent object SHALL NOT imply support for all descendants. Only explicitly identified non-visual metadata is exempt from visual assessment.

#### Scenario: Unsupported field on a familiar element

- **WHEN** an otherwise known element supplies a crop, mask, transform, rich-text, geometry, table, connector or chart field that the renderer does not honor
- **THEN** that field has a traceable limitation and the affected element/page cannot be graded fully supported

#### Scenario: Inherited and newly introduced state

- **WHEN** a page depends on unresolved theme, style, master/layout or component state, or a new visual descendant has no mapping
- **THEN** the actual owning path is reported and successful compilation does not promote the preview to full support

### Requirement: Limitations are precisely addressable

Each limitation SHALL include the relevant page and element identity where applicable, an exact input path, stable reason, support status, severity, bounded value summary and actionable review/remediation guidance. Program/page limitations SHALL use their true ownership path. Reports SHALL NOT expand source or binary payload contents as diagnostic text.

#### Scenario: Identical element IDs on separate pages

- **WHEN** two pages contain a limited element with the same local ID, including nested elements
- **THEN** their diagnostics remain distinguishable by page and exact path in the result and published receipt

#### Scenario: Unresolved visual bounds

- **WHEN** transforms, nested coordinates, text overflow, shadows or masks prevent a complete bounds check
- **THEN** the unresolved scope is reported rather than presenting the top-level frame test as complete geometry review

### Requirement: Support aggregation follows assessed state

Element, page and whole-result support SHALL aggregate conservatively using `unavailable > opaque > partial > supported`. Informational supported diagnostics SHALL NOT downgrade a supported state. Absence of diagnostics SHALL NOT alone prove support for unassessed state. File production success SHALL NOT upgrade visual support.

#### Scenario: Mixed assessment states

- **WHEN** a page contains supported and opaque/partial/unavailable states, or only supported informational diagnostics
- **THEN** its aggregate follows the declared ordering in both in-memory and persisted output, independent of diagnostic count or order

### Requirement: Factual errors fail a separate reliability gate

Preview SHALL expose a reliability status of `failed`, `requires-review` or `passed`, separate from support and publication completion. Known contradictions in data, geometry, topology, missing observations, axes or visibility SHALL fail the gate. Partial, opaque or unclassified visual state SHALL require review when no error already fails it. A passed automated gate SHALL NOT claim human visual review or editing fidelity.

#### Scenario: Known factual distortion despite produced files

- **WHEN** rendering drops a text-bearing shape's geometry, changes connector topology, coerces or bridges unknown observations, misrepresents proportions/hierarchy/flows/cumulative totals, loses numeric/OHLC/size channels, or contradicts declared axes/transforms/visibility
- **THEN** the reliability gate fails with the responsible paths, even if SVG and PNG publication completes successfully

#### Scenario: Appearance cannot clear the gate

- **WHEN** a page with a factual error receives favorable visual styling or successfully passes compilation/publication
- **THEN** the reliability status remains failed and its violations remain present

### Requirement: Standalone preview images communicate limitations

A page with failed reliability SHALL visibly identify itself as an unreliable preview; partial/opaque or otherwise review-required pages SHALL visibly indicate that review is needed. The indication SHALL identify diagnostic evidence without changing input state, original element IDs or canvas dimensions. Both the SVG and its PNG SHALL carry the indication.

#### Scenario: Image is inspected without its manifest

- **WHEN** a reviewer opens an image from a limited or factually incorrect page without reading its receipt
- **THEN** the image itself communicates the limitation and does not silently resemble fully accepted output

### Requirement: Diagnostic evidence preserves publication and source contracts

The returned result and final receipt SHALL carry consistent page/global support and reliability evidence. Publication `ok` and output completion SHALL continue to describe files actually produced, not correctness acceptance. Publication failures SHALL remain unavailable/incomplete failures and SHALL also fail reliability. Diagnostic assessment SHALL NOT mutate source-bound/opaque inputs or load new specialist dependencies through root imports or in-memory SVG use.

#### Scenario: Source-bound preview completes publication

- **WHEN** preview publishes a source-bound input with unresolved or failed visual state
- **THEN** source/input bytes remain unchanged, artifact hashes describe the actual labeled images, the receipt preserves the failed/review-required gate, and successful publication does not become an edit-fidelity claim

#### Scenario: Raster dependency or artifact write fails

- **WHEN** the established output publisher cannot complete its request
- **THEN** existing failure codes, retained partial artifacts and truthful manifest behavior remain in force and the reliability evidence cannot report passed
