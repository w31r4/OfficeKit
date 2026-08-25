## ADDED Requirements

### Requirement: Frozen A/B/C authoring pilot
The repository SHALL define ten authoring tasks across business review,
academic defense, technical architecture, analytical reporting, and brand
launch. It SHALL run Grid-default A, unconstrained-principles B, and authoring-
compiler C from the same packed package, inputs, model configuration, time
limit, and token limit for two fresh-context trials per task.

#### Scenario: Materialize a pilot task
- **WHEN** one task is prepared for A, B, and C
- **THEN** only the authoring-route instruction differs and hidden grader facts remain unavailable to every Agent

### Requirement: Independent deterministic oracles
Every run SHALL record package creation, import, verify, layout/design review,
source protection, plan/commit/resume state, editability, time, tokens, retries,
and a content/evidence oracle. Run identity and artifact hashes SHALL be stable
and auditable.

#### Scenario: Grade a generated deck
- **WHEN** an Agent exits successfully but its PPTX cannot reimport or its evidence is wrong
- **THEN** the task is a hard failure regardless of visual preference

### Requirement: Blind quality comparison
The pilot SHALL randomize arm identity and order for independent Codex visual
pairwise review and SHALL emit human spot-check material for key slides. It
SHALL score narrative, topic/style fit, text naturalness, silhouette and density
diversity, information relationships, and follow-up improvement separately
from hard validity.

#### Scenario: Judge two candidates
- **WHEN** a blind reviewer compares candidates
- **THEN** neither file name, prompt, montage label, nor metadata reveals its arm

### Requirement: Evidence-gated default switch
Compiler route C SHALL replace Grid-default only if hard-pass rate is at least
95% and within two percentage points of A, blind wins are at least 60% over A
and 55% over B, selected continuation success is at least 90%, and median time
and token use are no more than 1.5 times A. The frozen pilot is the first
decision packet; a post-fix rerun may be used only when it preserves that
packet, uses fresh blind judging, adds an independent unseen holdout, and
records the new rollout decision separately.

#### Scenario: One threshold fails
- **WHEN** any declared rollout threshold is not met
- **THEN** Grid remains the shipped default, C remains experimental, and documentation reports the failed metric without claiming completion

#### Scenario: All thresholds pass
- **WHEN** all thresholds pass on the frozen pilot or on a qualifying post-fix rerun plus unseen holdout
- **THEN** the Skill changes its no-direction route to C, a thirty-task expansion becomes part of the Presentation slow gate, and the historical pilot remains append-only evidence
