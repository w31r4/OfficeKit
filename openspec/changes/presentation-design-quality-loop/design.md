## Context

OfficeKit 0.9 made communication job, scenario, design source, selected direction,
and per-deck design grammar explicit. The resulting financial, management, and
brand dogfood decks are visibly different, but they also reveal the next gap:
principles do not guarantee that an Agent makes concrete page decisions at the
right time.

The current seven scenario guides are concise, generally 47 or 48 lines. They
state audience jobs, narrative patterns, visual carriers, and failure modes, but
do not yet provide the richer page grammar, chart/table de-defaulting, imagery
planning, and position discipline present in mature presentation workflows.
The recent financial and management examples still allowed generic white chart
surfaces to sit apart from their page grammar.

Public PPT Master materials reinforce that stable quality comes from routed
workflow authority, a durable design specification, representative previews,
separate quality review, and recoverable project state. The local Kimi research
materials likewise show the value of scenario-specific page grammar and explicit
anti-default instructions. OfficeKit will use these as observations only and
write its own clean-room guidance around the capabilities it actually ships.

## Goals / Non-Goals

**Goals:**

- Translate a chosen direction into concrete, repeatable decisions before the
  Agent builds the whole deck.
- Improve charts, tables, diagrams, imagery, density rhythm, geometry, and page
  variety without imposing a house style.
- Catch a weak design grammar on three representative pages rather than after a
  complete 14-page deck.
- Keep the one-sentence user experience: internal calibration must not become a
  design questionnaire.
- Improve generalization through process and independent evidence, not by tuning
  instructions against known sample outputs.

**Non-Goals:**

- Add another template library, visual DSL, design score, or layout recommender.
- Copy Kimi application text, private templates, PPT Master source, or fixed
  example designs.
- Require image generation, multimodal vision, subagents, or a particular host.
- Guarantee identical quality across models or claim that warning-free output is
  beautiful.
- Rerun the A/B/C matrix, add a permanent benchmark harness, or expand release
  gates.

## Decisions

### Deepen only the selected scenario

Each scenario guide will use the same clean-room decision structure:

1. audience job and evidence hierarchy;
2. narrative and density rhythm;
3. representative page archetypes;
4. visual-carrier selection;
5. chart, table, diagram, and image treatment;
6. geometry, line, surface, type, and spacing rules;
7. scenario-specific failure and repair questions.

The router still loads one primary guide and at most one genuine secondary
guide. This yields useful detail without placing all seven style documents in
the model context.

The guides specify relationships and ranges, not default HEX values, font names,
or layout IDs. User templates and brand systems remain authoritative.

### Calibrate with three representative pages

After writing the authoring plan and grammar, the C route selects:

- the opening or direction-setting page;
- an evidence/data page;
- the densest, most relational, or otherwise highest-risk page.

It composes and renders those pages first. The Agent reviews them together at
contact-sheet scale and inspects the highest-risk page at readable scale. It may
update the same authoring plan's design grammar before completing the deck.

This is an internal authoring checkpoint. It does not ask the user to select a
layout, and it does not publish a partial deck. Decks of four pages or fewer use
the full deck as the calibration spread.

### Make every visual carrier name its source strategy

The existing string `compositionIntent` remains unchanged. The Skill requires
it to state both the carrier and where that carrier comes from:

- user-provided asset;
- template or style-reference asset;
- sourced image with provenance;
- generated image;
- native chart, diagram, table, typography, or vector composition.

If image capabilities are unavailable, the Agent chooses a truthful native
carrier or asks for a core missing asset. It does not fill the page with unrelated
ornament or a fake image placeholder.

### Add decision recipes, not page templates

Shipped examples will demonstrate:

- de-defaulting chart surfaces, axes, grid, series, labels, and sources;
- using lines, bands, rules, and asymmetry rather than universal containers;
- integrating text with diagrams and images;
- using a recurring motif without making it the main carrier;
- composing one page from a clear information relationship.

Examples use only public OfficeKit primitives and explain what decision they
embody. They are not registered as layouts and do not prescribe one palette.

### Separate generation evidence from review judgment

The authoring route owns direction, grammar, and composition. The review route
reopens the plan, candidate, and rendered evidence, then judges whether the
declared carrier, hierarchy, rhythm, and constraints are visible. It does not
treat the generator's rationale as proof.

This separation can run in the same Agent context, a fresh context, or a host
with delegated review. The Skill describes the information boundary rather than
requiring subagents.

### Repair the failed layer instead of adding decoration

When calibration or final review finds a weak page, repair proceeds in order:

```text
claim and evidence
-> visual carrier
-> composition and hierarchy
-> styling and motif
-> motion
```

An Agent does not repair weak structure by adding circles, shadows, icons, or
animations. Chart problems are repaired inside the chart and its explanatory
relationship before adding surrounding decoration.

### Keep generalization evidence small and unseen

Implementation uses three new briefs that were not used to write the guidance:
one information-dense reader deck, one live narrative deck, and one
template/style-authority deck. Each runs once. Evidence records concrete failures
and whether the selected direction is visible; it does not assign an aesthetic
score or tune rules to reproduce a reference screenshot.

Only a repeatable product failure earns a minimal regression assertion. These
runs remain manual dogfood, not slow-gate inputs.

## Risks / Trade-offs

- [Scenario guides become prompt bloat] -> Keep one common structure, load only
  the selected guide, and move API detail to Help/examples.
- [Detailed guidance becomes a hidden template] -> Specify purpose,
  relationships, ranges, and counterexamples; prohibit fixed palettes, layout
  IDs, and mandatory page sequences.
- [Calibration increases latency] -> Limit it to three representative pages and
  reuse them in the final deck; short decks calibrate once as a whole.
- [The Agent rationalizes a weak page] -> Review rendered evidence separately
  and keep heuristic findings as warnings rather than accepting prose intent.
- [Image tooling is unavailable] -> Select an informative native carrier or ask
  for a core asset; never claim a decorative placeholder satisfies the plan.
- [More examples cause imitation] -> Annotate examples by decision and expose
  multiple visual grammars; do not register them as default templates.
- [Model quality remains variable] -> Define a stable process and quality floor,
  report limitations honestly, and avoid promising deterministic taste.

## Migration Plan

1. Expand the shared scenario-guide structure and selected-scenario routing.
2. Add the three-page calibration sequence to create and template/style-transfer
   routes using the existing task and authoring plan.
3. Add public-API decision examples and chart/image/diagram guidance.
4. Tighten the review-deliver route around independent evidence and the repair
   ladder without adding new review APIs.
5. Run three unseen dogfoods once, record concrete failures, and add only the
   smallest justified regression.
6. Update package inventory, coverage, and release notes.

Rollback restores the shorter guides and direct full-deck composition route.
Existing plans, tasks, PPTX files, APIs, and wire messages remain valid.

## Open Questions

- Whether calibration should always include the opening page or select the three
  highest-risk pages when the opening is intentionally minimal. Initial
  implementation should prefer opening + evidence + highest risk and revise only
  if real dogfood shows the opening contributes no calibration value.
