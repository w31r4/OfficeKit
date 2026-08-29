# Motion

Motion explains order, causality, comparison, focus, or continuity. Complete
the static composition first; animation cannot repair an empty or incoherent
page.

PPJ stores transitions in `pages[].transition` and semantic object animations
in `pages[].animations`. Targets are stable element IDs. Supported effects are
fade, wipe, fly, zoom, and pulse; start modes are `withPrevious`,
`afterPrevious`, and `onClick`. Text may build as whole or paragraph. Charts
may build all at once, by series, category, or element. Morph uses explicit
adjacent-page pairs.

## Delivery policy

- `reader`: static by default; animation requires explicit user intent.
- `hybrid`: one restrained base transition; animate only data, sequence, or
  causal relationships that gain meaning from reveal.
- `live`: semantic choreography is allowed, usually no more than four motion
  units on a page.

Never auto-advance or add sound unless the user explicitly requests it. Keep
one base transition and at most one section transition across a deck. High-noise
effects are not automatic choices.

## Six recipes

### Data Rise

Reveal bars, line points, or a small number of categories in data order. Keep
axes, labels needed for orientation, and sources visible. Do not use stagger to
hide an unreadable chart.

### Causal Reveal

Reveal causes, connectors, and effects in the direction of reasoning. A
connector appears with or after the node whose relationship it explains.

### Comparison Beat

Bring comparable objects in together, then reveal the conclusion. Preserve a
common scale and avoid making one side look stronger through timing alone.

### Focus Pulse

Pulse one already visible number, threshold, or risk. Repeated pulse becomes
noise and should trigger a review warning.

### Calm Continuity

Use short fade, wipe, or push transitions to maintain chapter rhythm without
turning page changes into the subject.

### Morph Continuity

Use adjacent pages and explicit one-to-one compatible object pairs to move,
scale, crop, or focus a repeated object. Charts do not participate in Morph;
use chart build or cross-fade. Pair keys and object identities must remain
unique and recoverable after re-import.

## Review

Check target existence, trigger order, delay/stagger, text/chart build type,
timing budgets, reader-mode authorization, and Morph adjacency. Then observe
playback in an available host. `structural` proves only that canonical timing
was written; report Keynote or PowerPoint evidence only when actually played.
