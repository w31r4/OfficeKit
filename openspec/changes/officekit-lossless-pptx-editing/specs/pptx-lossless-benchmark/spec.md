## ADDED Requirements

### Requirement: Benchmarks identify immutable real sources
The benchmark SHALL identify each external PPTX by SHA-256 and record its OPC inventory, relationships, master/layout/theme bindings, native object counts, text nodes, render evidence, editable node index, and declared edit targets without redistributing the source file.

#### Scenario: Benchmark freeze
- **WHEN** a benchmark manifest is generated for a source asset
- **THEN** every recorded inventory and target is tied to the exact source SHA-256

### Requirement: Repeated edits are deterministic
Each matrix operation SHALL run at least three times from a clean source copy and produce identical output hashes and mutation footprints.

#### Scenario: Three clean repetitions
- **WHEN** a text, grouped text, geometry, image, chart, or native-leaf matrix edit is repeated three times
- **THEN** all accepted runs have the same output hash, footprint, and independent package-oracle result

#### Scenario: Supplemental SmartArt canary
- **WHEN** the three immutable external decks contain no SmartArt
- **THEN** a repository-owned real closed-diagram package may provide the SmartArt leaf target only when the manifest and evidence label it as supplemental rather than third-party coverage

### Requirement: Completion requires independent package, visual, host, and Agent evidence
The capability SHALL remain incomplete until all benchmark no-ops, declared edits, second imports, package diffs, non-target renders, Windows PowerPoint checks, clean-install execution, full gates, hosted CI, and three independent Agent workflows pass.

#### Scenario: Code and unit tests pass but host evidence is missing
- **WHEN** implementation and automated tests pass but Windows PowerPoint or Agent 3/3 evidence is absent
- **THEN** the benchmark remains incomplete and reports the missing acceptance evidence

#### Scenario: Final acceptance
- **WHEN** every completion gate passes on the immutable sources and published package candidate
- **THEN** the benchmark report can mark the lossless-editing goal complete

### Requirement: Competitors are non-blocking controls
Kimi, HTML, and PPTD outputs SHALL be recorded when available but SHALL NOT replace source-PPTX fidelity requirements or block OfficeKit evaluation when unavailable.

#### Scenario: Competitor unavailable
- **WHEN** a comparison tool cannot run or its result cannot be reproduced
- **THEN** OfficeKit acceptance continues against the original PPTX and records the missing control
