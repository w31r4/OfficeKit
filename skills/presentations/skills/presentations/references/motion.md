# Presentation motion

Motion is a small communication layer over the editable slide model. It should
answer “what should the audience notice next?” rather than decorate every
object.

## Pick a motion mode

- **Speaking:** reveal a causal chain, chart series, or one important number;
  use `onClick` for deliberate beats and `withPrevious` for a single beat.
- **Reading:** keep the complete argument visible; use no object animation or
  only a restrained emphasis pulse.
- **Hybrid:** animate only the transition from evidence to conclusion. Do not
  make the reader wait for ordinary labels or footnotes.

Start with one or two animations on a page. Prefer the least surprising
effect: `fade` for appearance, `wipe` for data growth or directional flow,
`fly` for a deliberate entrance, `zoom` for a focal transition, and `pulse`
for emphasis. Use `slide.setMorph()` only when the same named object is
intentionally carried from one slide to the next.

## Typed surface

```js
slide.animations.add(target, {
  effect: "wipe",
  direction: "up",
  chartBuild: "series",
  start: "onClick",
  durationMs: 650,
});
slide.setMorph({
  durationMs: 700,
  pairs: [{ key: "hero", fromId: "overview-hero", toId: "detail-hero" }],
});
```

`textBuild` is `whole` or `paragraph`. `chartBuild` is `allAtOnce`, `series`,
`category`, `seriesElement`, or `categoryElement`. `start` is
`withPrevious`, `afterPrevious`, or `onClick`. The API is typed and bounded;
it is not a raw PresentationML timing editor.

## Review

After export, inspect the slide timing records and re-open the PPTX. Render the
changed pages for static layout checks, then verify that the intended sequence
is represented. Unknown imported `p:timing` and Morph extension graphs remain
opaque; do not flatten or replace them to make a new effect fit. A playback
check in a native Office host is separate evidence from package and layout
verification.
