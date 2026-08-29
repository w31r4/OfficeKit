## ADDED Requirements

### Requirement: Standalone PPJ command family
The package SHALL provide `officekit ppj import`, `inspect`, `check`, `build`, `render`, and `review`, and every command SHALL work without creating or loading a Task.

#### Scenario: Standalone authored workflow
- **WHEN** an Agent checks, builds, renders, and reviews a valid local PPJ without `--task`
- **THEN** the commands complete using only the declared program, assets, output paths, and local runtime

### Requirement: Direct file editing workflow
The CLI SHALL treat the PPJ file as the editable source and SHALL NOT require JavaScript authoring methods, a live REPL object, or a hidden heap.

#### Scenario: Agent edits JSON directly
- **WHEN** an Agent changes valid PPJ fields with an ordinary file editor and runs `check`
- **THEN** the checker validates the new file and reports its canonical revision without requiring an edit command log

### Requirement: Non-mutating fuzzy inspection
`officekit ppj inspect` SHALL support page, type, role, text, and fuzzy query filters, SHALL return zero or more stable IDs, and SHALL never select or mutate an object implicitly.

#### Scenario: Query returns multiple candidates
- **WHEN** several elements match a fuzzy text or role query
- **THEN** inspect returns every candidate with page, type, frame, summary, and capability information

### Requirement: Separate check, build, render, and review costs
The CLI SHALL keep structural validation, native build, page rendering, and artifact review as distinct operations. Build SHALL perform required structural validation but SHALL NOT automatically render or review.

#### Scenario: Fast iterative check
- **WHEN** an Agent runs `check` after a PPJ text edit
- **THEN** no PPTX, renderer, review provider, or Office Live bridge is loaded

### Requirement: Deterministic check repairs
`check --fix` SHALL perform only deterministic syntax formatting, canonical defaults, and safe path normalization and SHALL NOT alter content, layout, design, or source-bound intent.

#### Scenario: Semantic repair is unavailable
- **WHEN** a program contains an overlap, unsupported native mutation, or weak design decision
- **THEN** `--fix` reports the issue and leaves the semantic content unchanged

### Requirement: Optional durable Task binding
Successful `check` or `build` with `--task` SHALL save an immutable PPJ revision, hash, receipt, and relevant review/artifact binding; without `--task` no Task state SHALL be read or written.

#### Scenario: Fresh context resumes valid PPJ
- **WHEN** a task has a reviewed PPJ revision and a fresh context resumes it
- **THEN** OfficeKit materializes that revision and reports the PPJ path, hash, artifact, review state, and next action without restoring a JavaScript heap

### Requirement: Legacy plan tasks fail explicitly
Presentation tasks whose only durable authoring source is legacy `ctx.plan` SHALL remain listable but SHALL NOT be migrated or resumed as PPJ.

#### Scenario: Legacy task resume
- **WHEN** a user attempts to resume a legacy presentation-plan task under OfficeKit 2.0
- **THEN** the CLI reports an unsupported task schema and leaves every legacy file unchanged

### Requirement: Input artifacts are never overwritten
Import, build, render, and review SHALL reject output paths that alias a declared PPJ, source PPTX, or asset input.

#### Scenario: Output equals source path
- **WHEN** build output resolves to the source PPTX path
- **THEN** the command fails before invoking the compiler
