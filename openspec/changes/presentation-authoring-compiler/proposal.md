## Why

OfficeKit already exposes flexible Presentation authoring, imported-template,
source-derived reuse, review, and durable-task primitives, but its packaged
Skill routes an unspecified visual request into one Grid design system and
spreads advanced operations across a 1,000-line entrypoint. A fresh Agent can
therefore miss capable APIs or produce a stable but visually repetitive deck
even though the runtime can support a freer result.

## What Changes

- Add one durable, hash-bound authoring-plan slot to task REPL sessions so a
  fresh Agent can recover the communication brief, narrative, per-deck design
  grammar, ordered page plan, editorial rules, and next recipe.
- Rebuild net-new Presentation authoring around a deck-specific design grammar
  and free Compose primitives. Keep Grid Layout as an explicit scaffold rather
  than the implicit no-direction default.
- Replace the monolithic Presentations Skill surface with five progressively
  loaded task routes, six design-mechanism references, text-editing guidance,
  and design-review guidance.
- Extend Help metadata so high-value Presentation APIs are searchable by
  intent and linked to runnable recipes and required review.
- Bind the active authoring-plan revision into Presentation review and task
  commits, and report deterministic design invariants separately from
  non-blocking aesthetic heuristics.
- Add packed clean-install A/B/C Agent evaluation. Change the default route
  only if the declared quality, reliability, continuation, and cost thresholds
  pass.

## Capabilities

### New Capabilities

- `durable-authoring-plan`: Versioned task-plan storage, recovery, review
  binding, commit binding, migration, and publication safety.
- `presentation-authoring-workflow`: One-prompt brief-to-PPTX authoring with
  per-deck design grammar, progressive task routes, editorial guidance, and
  deterministic review.
- `presentation-agent-adoption`: Intent-indexed Help metadata and recipe
  coverage for the high-value Presentation API surface.
- `presentation-authoring-evaluation`: Reproducible packed A/B/C generation,
  continuation, blind-review, and rollout evidence.

### Modified Capabilities

None. No canonical repository spec currently defines these behaviors.

## Impact

- Additive REPL API: `ctx.plan([value], options)`; task manifest schema advances
  to 2 and REPL transport advances to 3 with read compatibility for v1 tasks.
- Additive `reviewArtifact` options and Presentation review evidence.
- Presentations Skill, runtime Help metadata, generated API documentation,
  task/review tests, clean-install packaging, and slow Agent evaluation change.
- Target release is `0.7.0`. Office wire protocol remains 2; C# Codec, PDF,
  Documents, Spreadsheets, Live adapters, and template asset formats do not
  change.
