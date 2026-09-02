## ADDED Requirements

### Requirement: Frozen dual-arm presentation Skill experiment

The experiment SHALL compare exactly two clean-room Presentation Skill overlays
over the same PPJ/compiler/runtime capability pack: a Shared What / What-kind /
How route and a concise scenario-first route with a task-local style brief.

#### Scenario: Arm identity is reproducible

- **WHEN** an experiment is prepared
- **THEN** each arm has a complete file tree, a SHA-256 manifest, the same
  common invariants, the same PPJ/capability references, and no Kimi source
  text, code, private template, or asset copied into either arm

#### Scenario: Production behavior remains unchanged

- **WHEN** the experiment branch is compared with its `origin/main` parent
- **THEN** no PPJ schema, Office wire message, codec, public API, production
  default route, or shipped template behavior is changed by the experiment

### Requirement: Scenario and capability coverage

The experiment SHALL freeze six communication scenarios and SHALL maintain a
separate capability ledger proving that scenario count does not stand in for
primitive coverage.

#### Scenario: Six scenarios are represented

- **WHEN** the case manifest is frozen
- **THEN** it contains analysis-decision, management-report,
  technical-engineering, academic-research, education-training, and
  brand-creative cases, each with audience, communication job, delivery mode,
  evidence boundary, and expected carrier

#### Scenario: Primitive coverage is explicit

- **WHEN** the capability ledger is checked
- **THEN** text, rich text, image/background, mask/opacity, line/connector,
  chart, table, group/z-order, formula, motion, and source-bound editing are
  each exercised by at least one case or explicitly marked unavailable

### Requirement: Paired 0-to-1 creation and 1-to-10 editing

The runner SHALL execute one 0-to-1 and one 1-to-10 case per scenario for each
arm, using identical frozen inputs and a seeded randomized arm order.

#### Scenario: Dense 0-to-1 page

- **WHEN** a 0-to-1 case is run
- **THEN** the Agent produces one high-density representative page with a
  primary carrier, three to six related content/evidence units, no supplied
  page skeleton, and a PPJ check/build/render/review record

#### Scenario: Two-step 1-to-10 edit

- **WHEN** a 1-to-10 case is run
- **THEN** the Agent imports the assigned complex source, edits one target page
  through exactly two serial steps (semantic first, visual/delivery second),
  exports the full source package, and records intermediate check/build plus
  final render/review/reimport evidence

#### Scenario: Non-target source preservation

- **WHEN** the final 1-to-10 output is audited
- **THEN** the target mutation, source hash, stable IDs, opaque content,
  non-target slide/package parts, relationships, and requested capability are
  independently verified; unsupported or stale mutations fail closed

### Requirement: Shared image-sourcing treatment

Both arms SHALL use the same narrow image-sourcing Skill, providers, rights
policy, candidate and download limits, and no-image fallback, while allowing
the Agent to choose queries and images independently.

#### Scenario: Image query is part of the measured behavior

- **WHEN** a task needs a photo, illustration, or icon
- **THEN** each arm writes its own query, chooses candidates and crop, and the
  runner records queries, provider responses, selection, asset hash, rights,
  timestamps, calls, and fallback decisions

#### Scenario: Provider failure is visible

- **WHEN** a provider or network request fails
- **THEN** the failure is reported as an incomparable pair or an explicit
  incomplete case, never silently replaced with a different source

### Requirement: Hard-gated structural and visual evidence

The experiment SHALL treat artifact validity, evidence honesty, source-bound
scope, layout integrity, and renderability as separate hard gates from
subjective quality.

#### Scenario: Invalid output cannot win on appearance

- **WHEN** an output fails schema/check, build/render, stable identity,
  source-bound scope, rights, evidence honesty, overflow, or occlusion checks
- **THEN** its task is marked failed regardless of blind visual score

#### Scenario: Chart and layered evidence stay visible

- **WHEN** a page contains bars, lines, labels, markers, error bars, masks,
  captions, or background images
- **THEN** the oracle checks that evidence-bearing objects are not hidden,
  cropped, or made unreadable by another layer

### Requirement: Multi-dimensional blind scoring and efficiency ledger

The experiment SHALL use a fixed 100-point quality rubric, two independent
anonymous Luna Max judging rounds, a small human calibration sample, and a
separate efficiency ledger.

#### Scenario: Anchored quality score

- **WHEN** an anonymized output is judged
- **THEN** the judge scores communication/evidence 20, hierarchy/readability/
  occupancy 15, design craft/creative specificity 20,
  functional/editability 15, layer/layout/render robustness 15,
  display/medium fit 10, and completion/polish 5 using fixed 1-to-5 anchors

#### Scenario: Blind paired preference

- **WHEN** a judge receives both outputs for one case
- **THEN** arm identity and run order are hidden, page order is randomized, the
  judge scores independently before giving win/tie/loss, confidence, and a
  concise reason

#### Scenario: Efficiency is not hidden in quality

- **WHEN** a run completes
- **THEN** the ledger records input/output/reasoning tokens when available,
  wall time, Skill file reads, tool calls, image searches, retries,
  render/review passes, and quality-per-cost ratios without adding cost to the
  100-point quality score

### Requirement: Scope-limited paired analysis

The final report SHALL analyze the frozen 12-task sample without claiming
universal generalization or changing the production default.

#### Scenario: Paired result summary

- **WHEN** both arms have valid results
- **THEN** the report includes task-level win/tie/loss, overall and per-route
  means/medians, per-scenario deltas, exact paired permutation or sign tests,
  bootstrap intervals, judge agreement, image/non-image strata, and raw
  efficiency metrics

#### Scenario: Research-only rollout

- **WHEN** the report is published
- **THEN** it states the sample and platform boundaries, preserves prior Skill
  invariants, and makes no automatic route, default, release, or production
  merge decision
