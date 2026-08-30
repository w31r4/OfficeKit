## Context

The Presentation wire and `PptxTransitionCodec` already define a canonical,
validated transition profile for all base effects. PPJ currently collapses
that profile to a small subset, causing authored loss and imported projection
loss before the native codec is reached.

## Goals / Non-Goals

**Goals:**

- Make PPJ faithfully express every existing canonical base transition.
- Use effect-specific validation so Agents receive actionable errors before
  native compilation.
- Preserve explicit zero and false values for timing and trigger properties.
- Reuse the existing source-binding proof and mutate only the target SlidePart.

**Non-Goals:**

- Custom transition extensions, sounds, media triggers, arbitrary preset IDs,
  or show playback automation.
- Treating Morph as a base effect or permitting source-bound Morph synthesis.
- Building a transition-effect test matrix already covered by the codec test.

## Decisions

### 1. PPJ mirrors canonical semantic state

`transition.type` accepts `none`, all 21 base effects, and `morph`. Optional
fields are presence-aware so explicit `false` and `0` survive projection and
rebuild.

### 2. One lowering profile owns defaults

A shared C# lowering helper applies the same defaults as the native transition
normalizer: medium speed, click advance enabled, effect-specific direction or
orientation, one wheel spoke, and no timed advance unless declared. Authored
and source-bound paths use that helper rather than maintaining two profiles.

### 3. Applicability is validated semantically

JSON Schema owns enums and numeric limits. The semantic validator rejects a
field that the selected effect cannot represent, such as `spokes` on fade or
`direction` on circle. This avoids silently accepting decorative state that
PowerPoint would discard.

### 4. Morph remains a distinct contract

Morph retains adjacent source-page and explicit object-pair validation. Base-
only fields are rejected. An omitted Morph duration continues to default to
800 ms.

### 5. Imported edits remain capability-bound

Projection exposes the complete base transition and a `setTransition`
capability only when the source slide is transition-editable or transition-
addable. A changed PPJ transition lowers back to the existing canonical wire;
the codec re-proves source hashes and fails closed for unsupported timing.

## Risks / Trade-offs

- [The PPJ validator and codec defaults drift] -> Centralize PPJ lowering and
  retain the existing codec-wide 21-effect test as the effect catalog oracle.
- [Third-party transitions contain unknown extension state] -> Project only
  the canonical base profile and do not issue edit capability for unknown
  timing topology.
- [Expanded vocabulary encourages noisy decks] -> The language remains
  complete while Motion guidance keeps high-noise effects explicit opt-in.

## Migration Plan

The expansion is additive. Existing PPJ remains valid and keeps its prior
defaults. Generated guidance is refreshed from the schema, and one existing
PPJ round-trip test gains a rich transition and one source-bound edit.

## Open Questions

None.
