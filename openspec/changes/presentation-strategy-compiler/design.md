## Context

The `presentation-motion-compiler` branch already provides authoring-plan v1,
four design sources, six design mechanisms, free Compose, durable tasks, plan-
bound review, and native motion. Its default path requires an abstract design
grammar but does not require a presentation scenario or a concrete selected
direction, and its most visible dogfood helper teaches one rounded-container
style. The repository also retains a common style guide that is not mandatory
in the new create route.

The change must improve the decisions a fresh Agent sees without copying local
Kimi application resources, introducing a template DSL, or growing another
evaluation harness.

## Goals / Non-Goals

**Goals:**

- Make the communication job, delivery context, scenario, source authority,
  direction, grammar, composition, motion, and review one ordered workflow.
- Publish a product doctrine and execute the same model through the Skill,
  task plan, resume state, and bounded review evidence.
- Provide concrete but open-ended scene craft while preserving per-deck visual
  freedom and native PowerPoint editability.
- Preserve authoring-plan v1 and existing task revisions.

**Non-Goals:**

- Copy or redistribute Kimi reference text, remote templates, or private
  implementation details.
- Add a new model call, template service, vector database, presentation codec,
  wire message, universal aesthetic score, or second review API.
- Replace user judgment, image understanding, source verification, or native
  playback evidence.
- Build another A/B/C, PromptBench, or cross-scenario matrix.

## Decisions

### Keep authoring-plan v1 additive

Add bounded fields inside the existing flexible `brief` and `design` objects.
The normalizer validates fields when present; the current create route writes
all strategy fields. Legacy plans remain byte-stable and receive a review
warning instead of an automatic rewrite. This avoids a task migration and a
second durable format.

### Keep scenarios and mechanisms orthogonal

The seven scenarios describe the audience job and organizational convention.
The six mechanisms describe how information behaves. A plan records one
primary scenario, at most one secondary scenario, and zero to two mechanisms.
Combining both avoids turning a scene guide into a fixed template.

### Use a clean-room doctrine and scenario corpus

The public doctrine is written from public presentation standards, observable
OfficeKit behavior, and general communication principles. Seven scenario files
share one OfficeKit-authored structure. No build or sync step reads the local
Kimi directory, and package tests reject its paths and identifiers.

### Use progressive disclosure

The main Skill remains under 350 lines and owns only routing and invariants. A
new create task reads the short doctrine, scenario policy, and exactly one
primary scene file. The existing common style guide becomes the shared visual
floor; long imported-editing and native-format details stay on-demand.

### Lock one direction without a default questionnaire

For self-directed work, the Agent internally considers two or three distinct
directions and persists only the chosen name and rationale. It asks the user
only when an unresolved choice changes design authority, audience, conclusion,
or evidence. Templates and explicit brand systems always win.

### Extend the current design review

`reviewArtifact` keeps its public call shape. The existing design report gains
strategy descriptors and bounded issue categories. It reuses the current
structural, layout, motion, delivery, and visual-review fields rather than
introducing six parallel reports. Facts remain an Agent/source task; aesthetic
heuristics remain warnings.

### Validate through focused real work

Modify existing contract tests only for the additive plan and review behavior.
Use one financial deck and two smaller management/brand decks as one-time
dogfood. Do not register those artifacts as slow gates or permanent golden
fixtures.

## Risks / Trade-offs

- [Scenario documents become prompt bloat] → Load only the router and selected
  scene, keep the shared floor concise, and link advanced material directly.
- [Heuristics overfit a few bad examples] → Detect only stable modeled patterns,
  report affected pages, and leave the final judgment to rendered review.
- [Plan v1 gains ambiguous optional fields] → Validate new fields when present,
  make them mandatory in the shipped create workflow, and report legacy plans
  explicitly.
- [The doctrine duplicates Skill guidance] → Keep the public document
  explanatory and the Skill references imperative; test only synchronized
  concepts and links, not byte identity.
- [A scene guide overrides a real template] → Encode and repeat the design-
  authority precedence in the router, template task, and review guidance.

## Migration Plan

1. Land the doctrine and OpenSpec artifacts without changing runtime behavior.
2. Add v1 strategy validation and task descriptors with legacy compatibility.
3. Add clean-room scenario references and route the create task through them.
4. Extend review and examples, then run the focused dogfood.
5. Bump the package to `0.9.0`, regenerate public docs, and perform one final
   repository/package verification pass.

Rollback removes the additive route and fields while existing legacy plans
continue to work; no persisted task migration or Office wire rollback is
required.

## Open Questions

None. Source policy, doctrine surface, scenario/mechanism relationship, and
medium-fit behavior were explicitly selected before implementation.
