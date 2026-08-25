## Context

The Presentation runtime already supports free Compose trees, native theme and
layout authoring, template design profiles, source-derived reuse, SVG edits,
post-edit review, and durable task commits. The packaged Skill nevertheless
selects one Grid layout library whenever the user gives no visual direction,
and its large entrypoint makes newly added APIs hard for a fresh Agent to find.

Durable tasks currently persist Office/PDF revisions, Edit Plans, reviews,
constraints, and next actions. They deliberately do not restore JavaScript
heap state and have no validated place for an authoring brief or design plan.

## Goals / Non-Goals

**Goals:**

- Recover one bounded, versioned authoring plan in a fresh task session.
- Let each deck own a design grammar instead of inheriting one default visual
  system.
- Keep creative decisions with the Agent while OfficeKit validates, compiles,
  reviews, and preserves evidence.
- Make high-value APIs discoverable by intent through a small Skill and Help.
- Use reproducible evidence before changing the default authoring route.

**Non-Goals:**

- A PPTD file format, universal presentation AST, or model call in runtime.
- A new PowerPoint codec operation or Office wire change.
- Deterministic claims about aesthetic quality.
- Copying Kimi prompts, templates, styles, or private implementation details.
- Changes to Documents, Spreadsheets, PDF, Live adapters, or template assets.

## Decisions

### One immutable task-plan stream

Add `ctx.plan([value], options)` instead of several brief/design/page-plan
methods. The first accepted schema is
`office-kit/presentation-authoring-plan/v1`. A plan is plain JSON, at most
256 KiB, and contains at most 64 ordered pages.

Writing canonicalizes JSON, hashes it, writes a private read-only
`plans/<sha256>.json`, and atomically advances the task's active descriptor.
`expectedSha256` is mandatory after an active plan exists. Equal content is
idempotent. Calling without a value reads and revalidates a defensive copy.

Alternatives rejected:

- Arbitrary task-local files are not discoverable or integrity-bound.
- A generic key/value memory surface would restore unbounded model state.
- A public DeckPlan compiler would freeze an unproven authoring IR.

### Active and reviewed plan bindings

Task schema 2 stores one active plan descriptor. Every successful commit also
snapshots that descriptor. A Presentation review supplied with an authoring
plan records the same hash. A commit fails when the active plan, review plan,
and candidate do not agree.

Changing a plan after a reviewed commit makes the plan `working`; publishing
is blocked until another reviewed commit binds it. This prevents an Agent from
changing intent and then publishing an artifact reviewed under the old intent.

Task schema 1 is normalized in memory as `plan: null`. Read-only listing does
not rewrite it; the first mutating task operation writes schema 2. REPL
protocol advances to 3 because the ready envelope and context API change.

### Authoring plan is orchestration, not artifact source

The plan records:

- communication brief and immutable evidence references;
- narrative sections and ordered page reader-tasks;
- design-source mode and zero to two mechanism packs;
- a concrete per-deck design grammar;
- editorial rules, unresolved decisions, recipe, and next action.

MJS/Compose remains the executable authoring source. Existing imported
`presentation.designProfile()` remains descriptive evidence and is not
renamed or overloaded as the source-free design grammar.

### Mechanisms generate a grammar; they are not styles

Six English mechanism references describe narrative choices, composition,
density rhythm, evidence treatment, and failure patterns. They contain no
mandatory palette, font, fixed coordinates, or layout IDs. The Agent writes
actual choices into the plan's design grammar for each deck.

Grid Layout remains a selectable scaffold. It is never an implicit fallback
for an unspecified visual direction. A template, brand system, or reference
artifact remains authoritative when supplied.

### Progressive Skill and Help adoption tiers

Keep the main Skill below 350 lines and route to five task documents. Detailed
native edit profiles remain in references. Help gains additive adoption
metadata. Every `golden` entry must link a task recipe, runnable example,
capability precondition, and review obligation; `advanced` and `compatibility`
entries remain searchable without being loaded into the default workflow.

Generated API documentation remains downstream of Help metadata. No parallel
JSDoc example catalog is introduced.

### Design review extends the existing review boundary

`reviewArtifact` accepts `authoringPlan` and `changedPageIds`. Its Presentation
report includes a `design` section bound to the plan hash.

Machine-provable plan invariants can fail review: invalid/mismatched plan,
page-count mismatch, required unresolved fields, strict palette/font breaches,
content-budget violations, or undeclared changed pages. Repeated silhouettes,
density rhythm, card-like repetition, weak hierarchy, and repeated title forms
are warnings. They never become an aesthetic conformance claim.

### Evidence-gated rollout

The old Grid route, an unconstrained freeform route, and the compiler route run
from the same packed package and inputs in fresh Codex contexts. The compiler
route replaces the default only when the fixed reliability, preference,
continuation, and cost thresholds pass. The frozen pilot remains append-only;
the approved post-fix rerun additionally requires fresh blind judging and an
independent unseen holdout. Failed evidence preserves Grid as the default;
successful evidence leaves Grid as an explicit fallback.

## Risks / Trade-offs

- [Plan prompts increase tokens] → Keep the schema compact, Skill progressive,
  and cost threshold at 1.5x the Grid median.
- [Task schema migration corrupts old work] → Normalize v1 read-only and test
  mutation-time v2 migration against frozen fixtures.
- [Design checks over-constrain creativity] → Fail only explicit invariants;
  report aesthetic heuristics as warnings.
- [A single mechanism catalog becomes another template system] → Forbid fixed
  tokens/layout IDs and require a concrete per-deck grammar.
- [Blind model review is self-referential] → Randomize identity/order, retain
  deterministic oracles, and require a human spot check of key pages.
- [Skill split hides advanced capabilities] → Keep one-level references and
  index every golden API through Help and recipes.

## Migration Plan

1. Add schema-2 task storage and protocol-3 REPL compatibility without changing
   existing task behavior when no plan exists.
2. Split the Skill while preserving the current Grid route; freeze A evidence.
3. Add the compiler route, design review, Help adoption metadata, and B/C
   experimental instructions.
4. Run the packed pilot and any qualifying post-fix holdout. Switch the default
   only after the declared thresholds pass, preserving every earlier packet.
5. Release as 0.7.0. Rollback may restore the prior Skill route while retaining
   readable plan files and schema-2 tasks.

## Open Questions

None. Public DeckPlan compilation remains explicitly deferred until the Agent
evidence justifies a separate change.
