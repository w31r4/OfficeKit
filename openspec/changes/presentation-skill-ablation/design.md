## Context

The production branch already exposes PPJ as the public Presentation language and
contains focused references for communication jobs, scenarios, visual attention,
fonts, shapes, charts, media, layers, motion, imported native references, and
review. It also contains a mature source-bound import surface. The open question
is not whether those primitives exist, but how an Agent should retrieve and
sequence the guidance when creating or editing a page.

This change is a research-only comparison. The repository's main worktree is
dirty, so all experiment files and runs live in an isolated worktree based on
the execution-time `origin/main`. No production Skill, PPJ schema, compiler,
wire message, template, or default route is changed by the experiment.

## Goals / Non-Goals

**Goals:**

- Compare a three-layer What / What-kind / How route with a concise,
  scenario-first, task-local-style-brief route under identical PPJ and runtime
  capabilities.
- Cover six communication scenarios in both one-page creation and complex
  source-bound local editing.
- Measure visual communication, design craft, functional fidelity, render and
  layer integrity, display fitness, completion, token use, time, and tool use.
- Make image-query choice part of the treatment while keeping the image tool,
  providers, rights rules, and query budgets common.
- Preserve and verify the existing hard rules for evidence, sparse data,
  occlusion, render repair, PPJ identity, opaque preservation, and rights.
- Produce an auditable report with paired results, blind judgments, human
  calibration, limits, and follow-up hypotheses.

**Non-Goals:**

- Select or merge a new default Presentation Skill.
- Change PPJ, JavaScript APIs, the codec, Office wire, image providers, or
  production templates.
- Claim that six scenarios generalize to all presentations, or that structural
  evidence proves PowerPoint playback.
- Build a reusable benchmark framework, add a permanent evaluation matrix, or
  tune either arm after seeing the frozen results.
- Enable host image generation in the primary image-search comparison.

## Decisions

### 1. Use two clean-room overlays over one shared capability pack

The experiment has two arm directories. The Shared arm exposes the same
information through explicit `what`, `what-kind`, and `how` layers. The Kimi-style
arm keeps a short router, selects one route and scenario, and writes a task-local
style brief before PPJ composition. Both arms reference the same PPJ and
capability files and receive a common invariants document; only the navigation,
retrieval order, and intermediate style-brief artifact differ.

This tests the end-to-end guidance format without accidentally testing a weaker
primitive set. The common manifest records hashes and rejects drift between arms.
The Kimi-style wording and examples are clean-room writing; no Kimi source text,
code, private template, or asset is copied.

### 2. Freeze cases and use paired author sessions

The case manifest contains six scenarios, six 0-to-1 page briefs, six complex
source-page edit briefs, source and asset hashes, design-source metadata,
expected capability tags, and acceptance contracts. Each case is run once per
arm in a seeded, randomized order. The 0-to-1 output is one dense representative
page. The 1-to-10 output is a full source package with exactly one target page
edited through two serial steps: semantic first, then visual/delivery repair.

The source fixture for each scenario is selected and hashed before any model run.
Both arms start from identical bytes. Intermediate checks happen after the first
edit; final render, review, reimport, and non-target package checks happen after
the second edit.

### 3. Treat image retrieval as part of Skill behavior

Both arms use the same narrow `officekit image` route, provider allowlist,
candidate limit, rights filter, download protection, and no-image fallback. The
Agent independently chooses English queries, candidates, crop, and placement.
The runner archives queries, provider responses, candidates, selected asset
hashes, rights metadata, and timestamps. It does not preselect a common image
pool. Provider/network errors are recorded as an incomparable pair rather than
silently switching sources.

The primary experiment does not expose image generation. A future image
generation comparison must be a separate frozen study.

### 4. Separate hard gates, quality, and efficiency

Hard gates cover schema/check, build/render, stable IDs, source-bound scope,
opaque and non-target preservation, evidence honesty, rights, overflow, and
occlusion. A hard-gate failure is a task failure even when the page is visually
attractive.

The 100-point quality score combines blind judgment and deterministic evidence:
communication/evidence 20, hierarchy/readability/occupancy 15, design craft
and creative specificity 20, functional/editability 15, layer/layout/render
robustness 15, display and medium fit 10, and completion/polish 5. Token,
wall-time, file-read, tool-call, image-search, retry, and render/review counts
remain separate efficiency metrics.

Two independent fresh Luna Max judge sessions score anonymized A/B outputs in
random order. A four-pair human calibration checks judge reliability but never
tunes the arms. Results use paired win/tie/loss, means and medians, exact paired
permutation/sign tests, and bootstrap intervals; the report explicitly treats
`n=12` as exploratory.

### 5. Keep the runner disposable and the evidence inspectable

One small runner prepares the isolated arm, invokes `codex exec` with
`gpt-5.6-luna` and `model_reasoning_effort=max`, captures JSONL usage and file
reads, invokes the PPJ checks, and writes an evidence ledger. It is not added to
the production CLI or package. Raw runs stay outside the repository or under an
ignored `runs/` directory; only the frozen manifests, rubric, methodology and
aggregate report are tracked.

## Risks / Trade-offs

- [Network search results drift] → Run paired arms in an interleaved seeded order,
  archive provider responses and timestamps, and report image results separately;
  do not claim a pure text-only causal effect.
- [The concise arm hides a rule in a summary] → Generate both overlays from the
  same common invariants and capability hashes; route smoke must prove every
  mandatory rule is reachable.
- [The three-layer arm pays more context tokens] → Record actual file reads and
  model usage, and report quality/cost frontiers rather than hiding the cost.
- [Aesthetic judge favors one style] → Blind the arm identity, require anchored
  dimension scores and pairwise confidence, and use human calibration.
- [Complex source edit mutates unrelated content] → Use source hashes, target
  capability tags, second import, non-target part hashes and an independent
  package oracle as hard evidence.
- [A visual result is mistaken for playback proof] → Label evidence as
  structural/render unless a real host playback record exists.
- [Post-hoc tuning overfits the 12 tasks] → Freeze manifests and Skill hashes
  before runs; any later revision starts a separately named experiment.

## Migration Plan

1. Create the isolated branch/worktree and OpenSpec artifacts.
2. Write the shared invariants/capability manifest and the two arm overlays.
3. Freeze the six scenarios, source fixtures, assets, edit operations, rubric and
   run seed; hash every input.
4. Run lightweight route/manifest checks and one real smoke per arm.
5. Run the 24 author sessions and the 24 anonymous Luna Max judge sessions; add
   four human calibration records.
6. Produce the paired analysis and scope-limited research report.
7. Review the report with the user. Do not merge an arm or change the production
   default as part of this change.

Rollback is deletion of the experiment worktree and ignored raw runs. The
production branch remains byte-for-byte unaffected by the experiment.

## Open Questions

None for this run. The six scenarios, two serial edits, independent image
queries, two blind judge rounds, human calibration, quality weights, model,
efficiency separation, and research-only rollout are fixed by the approved plan.
