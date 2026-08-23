## ADDED Requirements

### Requirement: Deterministic intent matrix

The platform SHALL define at least ten unique editing intents for each frozen
sample and SHALL run each intent three times from the original source bytes.
Each run SHALL use only the packed public OfficeKit API, SHALL second-import the
result, and SHALL retain its individual outcome.

#### Scenario: Three clean repetitions
- **WHEN** one declared intent is evaluated
- **THEN** three independent imports start from the pinned source SHA-256
- **AND** a deterministic pass requires identical output hashes and oracle records

### Requirement: Independent package and pixel oracle

The evaluator SHALL verify the source SHA-256 after execution, exact OPC part
sets or declared additions, exact non-target part bytes, relationship graphs,
masked target XML/SVG, nested packages, second import, and exact non-target
page pixels. The evaluator SHALL NOT accept an OfficeKit edit plan as proof.

#### Scenario: Runtime receipt disagrees with package bytes
- **WHEN** runtime metadata reports an allowed edit but the independent package or pixel check fails
- **THEN** the run fails with the independent mismatch as its reason

### Requirement: Codex continuation matrix

The platform SHALL define one whole continuation task per sample and SHALL run
each task three times in a fresh ephemeral Codex context. Every trial SHALL use
the installed Presentations Skill, create and resume a durable OfficeKit task,
produce two reviewed commits in separate REPL sessions, resume again, and
publish without overwriting the source or an existing output.

#### Scenario: Nine isolated continuation trials
- **WHEN** the continuation matrix runs
- **THEN** each of the three task definitions receives three isolated trials
- **AND** every trial records its Codex trace, durable task evidence, output, and oracle result

### Requirement: Fail-closed evidence

Raw OOXML/ZIP mutation, Python, HTML/PPTD, `@oai/artifact-tool`, blank-deck
reconstruction, silent fallback, source mutation, output overwrite, missing
review/commit/resume evidence, or any oracle failure SHALL fail that trial.
The harness SHALL preserve the unmodified failure reason and SHALL NOT weaken
an expected footprint or pixel mask to obtain a pass.

#### Scenario: Agent uses a forbidden path
- **WHEN** a trace or authored task script shows a forbidden implementation path
- **THEN** the trial fails even if a visually plausible PPTX was produced

#### Scenario: Product capability is unsupported
- **WHEN** the public package refuses a declared task or the output fails an oracle
- **THEN** the baseline records the original failure without modifying product code or acceptance rules
