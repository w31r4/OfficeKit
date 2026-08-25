# Presentation motion

Motion is the behavior layer of an already complete page. It answers what the
audience should notice next. It does not supply the page's visual interest.

## Choose the delivery mode

- `reader`: no automatic object animation; keep the complete argument visible.
- `hybrid`: one restrained transition plus animation for data, sequence, or
  causality only.
- `live`: deliberate speaking beats, normally no more than four motion units
  per page.

Use no automatic advance or sound. Keep at most one baseline transition and
one section transition across a deck. High-noise effects require an explicit
stage direction from the user.

## Typed surface

```js
slide.animations.add(target, {
  phase: "entrance",
  effect: "wipe",
  direction: "up",
  start: "afterPrevious",
  durationMs: 650,
  delayMs: 0,
  textBuild: "paragraph",
  chartBuild: "category-element",
  staggerMs: 90,
  animateChartBackground: false,
});
```

Effects are `fade`, `wipe`, `fly`, `zoom`, and `pulse`. Starts are
`withPrevious`, `afterPrevious`, and `onClick`. Text build is `whole` or
`paragraph`. Chart build is `all-at-once`, `series`, `category`,
`series-element`, or `category-element`.

Remove a single returned record with `slide.animations.remove(record)`, or
clear the canonical graph with `slide.animations.clear()`. Unknown imported
timing stays opaque and cannot be reconstructed through this surface.

The complete runnable example is
[`officekit-motion-workflow.mjs`](../examples/officekit-motion-workflow.mjs).

## Data Rise

Use when a chart's order is part of the claim. Wipe from the baseline or reveal
time categories in sequence; do not animate legends, footnotes, or decorative
backgrounds.

```js
slide.animations.add(chart, {
  effect: "wipe",
  direction: "up",
  chartBuild: "category-element",
  start: "onClick",
  durationMs: 650,
  staggerMs: 90,
  animateChartBackground: false,
});
```

## Causal Reveal

Use for a process or causal chain. Reveal each node, its outgoing connector,
then the conclusion. A connector never arrives before both endpoints exist.

```js
slide.animations.add(cause, { effect: "fade", start: "onClick" });
slide.animations.add(link, { effect: "wipe", direction: "right", start: "afterPrevious" });
slide.animations.add(effect, { effect: "fade", start: "afterPrevious" });
```

## Comparison Beat

Use for two alternatives that should be perceived together. Bring both sides
in on one beat, then reveal the decision.

```js
slide.animations.add(left, { effect: "fly", direction: "left", start: "onClick" });
slide.animations.add(right, { effect: "fly", direction: "right", start: "withPrevious" });
slide.animations.add(conclusion, { effect: "fade", start: "afterPrevious" });
```

## Focus Pulse

Use once on an object that is already visible. It is suitable for a threshold,
risk, or decisive number; it is not a page-entry effect.

```js
slide.animations.add(riskNumber, {
  phase: "emphasis",
  effect: "pulse",
  start: "afterPrevious",
  durationMs: 500,
});
```

## Calm Continuity

Use one short transition to keep chapters coherent. Prefer `fade` for a quiet
change and `push` when the reading direction matters.

```js
nextSlide.setTransition({
  effect: "fade",
  speed: "fast",
  durationMs: 450,
  advanceOnClick: true,
});
```

## Morph Continuity

Use only for the same semantic object on adjacent slides. Pass real source and
destination objects. Charts do not participate; use Data Rise instead.

```js
detailSlide.setMorph({
  from: overviewSlide,
  durationMs: 800,
  pairs: [{ key: "hero", from: overviewHero, to: detailHero }],
});
```

OfficeKit assigns the paired Selection Pane identity to both objects, rejects
non-adjacent slides, incompatible object kinds, duplicate objects, name
conflicts, charts, and conflicting transitions, and restores the pair after a
second import.

## Review

Inspect with `presentation.inspect({ kind: "animation,morph" })`. Reopen the
PPTX, render every changed page, and verify static composition before playback.
Record `playbackEvidence` as `structural`, `keynote`, or `powerpoint`.
Structural evidence proves the package graph, not host playback.
