## Context

The branch already has a bounded source-free timing writer, chart build records,
basic Morph metadata, and a short Skill route. It does not yet connect those
records to the durable authoring plan, the C authoring route, visual composition,
or post-edit review. The implementation must preserve existing plan readers,
unknown imported timing, and Office wire protocol 2.

## Goals / Non-Goals

**Goals:**

- Make delivery mode, visual carrier, composition, and motion intent one durable
  authoring path.
- Compile typed effects, text/chart builds, order, stagger, and true adjacent-slide
  Morph into canonical native structures.
- Give a fresh Agent concise design/motion recipes and a review report that can
  distinguish structural evidence from real host playback.
- Exercise the path with one financial deck, one causal architecture deck, and
  one Morph brand deck.

**Non-Goals:**

- No PPTD format, universal AST, model call, sound, media control, arbitrary
  motion path, or raw Office.js/XML escape hatch.
- No reconstruction of unknown imported timing or Morph graphs.
- No new benchmark matrix, repeated clean-install experiment, or Windows host
  claim in this change.

## Decisions

### Additive plan fields, legacy-compatible normalization

Keep the v1 plan schema and the existing string `compositionIntent`. The C route
uses that string to name the page's visual carrier and composition job; it does
not add a structured composition object. Plans may add optional `motionIntent`.
`deliveryMode` and `motionPolicy` default to `hybrid` and `adaptive` when absent;
the C Skill writes them explicitly. Plan validation bounds recipes, units,
transitions, and trigger values without parsing arbitrary composition prose.

### Semantic motion before native timing

The Agent chooses a page recipe and target roles. The existing object model maps
those roles to stable objects, then the codec lowers the result to one canonical
timing graph. Semantic animation count is capped at 32 and expanded timing nodes
at 64. Unsupported combinations fail before writing bytes; no field is silently
ignored.

### Native Morph pairing across adjacent slides

`setMorph({ from, pairs })` accepts objects from adjacent slides in the same
presentation. The compiler validates compatible non-chart objects, prefixes both
source and destination native names with `!!key`, and writes one canonical Morph
extension on the destination transition. Imported extension graphs remain
opaque unless they match this profile.

### Review is evidence, not taste automation

Motion review checks target identity, order, build compatibility, limits, plan
agreement, Morph adjacency, and reader-policy violations. Empty composition,
repetition, and excessive effects are warnings. `playbackEvidence` is explicitly
`structural`, `keynote`, or `powerpoint`; structural evidence never implies host
playback.

### C-route design guidance

The Skill uses a deck-specific grammar and six recipes inspired by public Kimi/PPT
Master observations: theme roles, page archetypes, visual carriers, density
rhythm, assets, and micro-decoration. It does not copy private prompts or force
a fixed palette/layout. Motion is selected only after composition has a dominant
visual carrier.

## Risks / Trade-offs

- [A composition string may be vague] → Keep the schema small and require the C
  Skill to name one visual carrier before motion; review warns when the declared
  carrier cannot be matched to the completed page.
- [Chart stagger cannot be represented for a graph] → Reject that combination
  with an explicit capability error instead of dropping the stagger.
- [Imported timing is richer than the canonical profile] → Preserve it opaque
  and expose capability false.
- [Visual richness is not fully machine-verifiable] → Use deterministic carrier
  and occupancy warnings, then report visual review honestly.
- [Host playback is unavailable on macOS] → Keep structural evidence separate
  and record Keynote/PowerPoint only when explicitly observed.

## Migration Plan

1. Add OpenSpec artifacts and additive authoring-plan validation.
2. Implement adaptive route, composition guidance, and motion intent descriptors.
3. Harden runtime limits, chart build fields, and cross-slide Morph lowering.
4. Add inspect/review evidence and Help/Skill recipes.
5. Generate the three real decks, record host evidence, then run the single final
   gate and publish metadata for 0.8.0.

## Open Questions

None. The public API remains typed and the existing task/repl and wire versions
remain compatible.
