## 1. Change baseline and contracts

- [x] 1.1 Validate the OpenSpec proposal, design, delta specs, and task graph in strict mode.
- [x] 1.2 Record the frozen A baseline at `origin/main@a0452867` and the conditional default-switch rule in the evaluation manifest.
- [x] 1.3 Add fast contract tests that keep the Office wire protocol at 2 and reject eager Office, MuPDF, provider, or Live runtime initialization.

## 2. Durable authoring plan

- [x] 2.1 Implement the canonical `office-kit/presentation-authoring-plan/v1` validator, size and page limits, stable JSON encoding, and SHA-256 descriptors.
- [x] 2.2 Add immutable private `plans/` storage, idempotent writes, exact `expectedSha256` stale-write protection, and artifact-reference validation.
- [x] 2.3 Upgrade task manifests to schema 2 with read-only schema-1 compatibility and mutate-time lossless migration.
- [x] 2.4 Bind each artifact commit to the active plan revision and block publication after the plan changes until a new reviewed artifact commit is made.
- [x] 2.5 Add task-store tests for plan validation, limits, hashes, atomic writes, idempotency, stale updates, migration, review binding, and publish blocking.

## 3. REPL and task user surface

- [x] 3.1 Upgrade the REPL protocol to 3 and add `ctx.plan()` read/write behavior without loading presentation runtimes.
- [x] 3.2 Expose compact plan descriptors through `ctx.task.plan` and `session.ready.task.plan` while keeping full plan bodies behind `ctx.plan()`.
- [x] 3.3 Extend `officekit tasks <id>` and resume summaries with mode, page count, recipe, plan state, hash, next action, and pending review information.
- [x] 3.4 Update REPL, task, clean-install, and lazy-root tests for schema-1 compatibility and schema-2 authoring plans.

## 4. Presentation authoring review

- [x] 4.1 Extend `reviewArtifact()` with optional `authoringPlan` and `changedPageIds` inputs and a deterministic `design` report section containing the plan hash.
- [x] 4.2 Implement blocking checks for page count, required unresolved items, strict color/font constraints, content budgets, and invalid changed-page scope.
- [x] 4.3 Implement bounded warnings for repeated composition, density jumps, card-wall patterns, repeated title syntax, and design drift without claiming aesthetic judgment.
- [x] 4.4 Verify local edits keep non-target pages stable, preserve honest `visualReview: unavailable`, and keep AnyDoc limited to declared content-coverage gaps.
- [x] 4.5 Add review tests for invariant failures, heuristic warnings, plan/commit revision mismatches, local-edit drift, and no-vision behavior.

## 5. Progressive Presentation Skill

- [x] 5.1 Reduce the Presentations `SKILL.md` to at most 350 lines containing routing, safety, four design sources, five golden paths, and capability search.
- [x] 5.2 Add `tasks/create.md`, `create-from-template.md`, `edit-existing.md`, `continue.md`, and `review-deliver.md` with deterministic workflows.
- [x] 5.3 Add one-level references for authoring plans, six design mechanism packs, audience-facing text editing, design review, and advanced imported editing.
- [x] 5.4 Encode the evidence-first editorial flow, page-local editing rule, explicit global redesign rule, and zero-to-two mechanism-pack constraint.
- [x] 5.5 Preserve Grid Layout as an explicit scaffold with no silent fallback; after the post-fix blind and unseen evidence cleared the rollout thresholds, switch the no-direction self-directed route to C.
- [x] 5.6 Update OfficeKit routing and REPL references so a fresh Agent can select create, template, existing-edit, continue, and review-deliver paths.
- [x] 5.7 Add Skill line-count, routing, mechanism, source-mode, editorial, portability, and reference-sync tests.

## 6. Help adoption index

- [x] 6.1 Extend Presentation Help records and search indexing with `adoptionTier`, `useWhen`, `avoidWhen`, `requires`, `review`, and `recipes`.
- [x] 6.2 Classify all public Presentation APIs as golden, advanced, or compatibility and add complete adoption metadata to every golden API.
- [x] 6.3 Add real recipe paths and minimal runnable examples for Compose, AutoLayout, layout validation, design profile, template planning, inspect, source reuse, component reuse, SVG editing, and layout placeholders.
- [x] 6.4 Regenerate API documentation from Help metadata and add natural-language adoption-search and broken-recipe tests.

## 7. Packed clean-install workflows

- [x] 7.1 Add a one-sentence self-directed creation scenario that writes a plan, composes a deck, reviews, commits, resumes, edits locally, and publishes.
- [x] 7.2 Add a template-conditioned creation scenario using design profile, template planning, source-slide/component reuse, review, and resume.
- [x] 7.3 Verify both scenarios from the packed artifact with no project-local OfficeKit dependency and no raw OOXML patching.

## 8. A/B/C authoring experiment

- [x] 8.1 Add a versioned pilot manifest for 10 tasks across five scenarios, two trials, three arms, fixed model/material/time/token conditions, and anonymous randomized judging.
- [x] 8.2 Implement the packed fresh-context runner, artifact oracle, metric collector, retry accounting, and independent visual-blind-review packet generator.
- [x] 8.3 Run all 60 pilot generations and publish raw run evidence, hard-gate metrics, fresh read-only Codex blind quality judgments, time, token, and retry summaries.
- [x] 8.4 Preserve the frozen pilot result: C scores `50%` over A (below the `60%` requirement) and `60%` over B (meeting the `55%` requirement), so the historical scorer records `keep-A` without being overwritten.
- [x] 8.5 Register the diverse 30-task expansion contract in the Presentation slow gate and verify the selected post-fix continuation evidence at `23/23` (`100%`), while keeping the full live 30-task Codex matrix explicitly pending rather than fabricating pass records.

## 9. Release and closure

- [ ] 9.1 Update coverage, release notes, package contents, README, and version metadata for the evidence-supported `0.7.0` surface.
- [ ] 9.2 Run affected fast gates after each atomic commit and then run full npm, docs, package, release, packed-install, and hosted CI gates at the milestone boundary.
- [x] 9.3 Confirm no Office Codec wire, C# codec, PDF, Spreadsheet, Document, Live Add-in, or template-asset-format changes entered the change.
- [ ] 9.4 Archive the OpenSpec change only after implementation tasks and evidence-gated rollout decisions are complete.
