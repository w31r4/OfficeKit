# Design mechanisms

Mechanisms guide how information behaves across a deck. They do not provide a
palette, font, coordinate set, or layout ID. Select zero to two, then write a
specific design grammar for the actual audience and content.

## Cross-cutting authoring floor

Mechanisms are deliberately open-ended, but a self-directed deck still needs a
small quality floor. Put the floor in the deck's `designGrammar`, rather than
silently inheriting a library style:

- state a readable body and caption size in the model's font-size units; shorten,
  split, or restructure before shrinking text;
- choose text/background pairs with clear contrast, and check them at contact-
  sheet scale instead of trusting a token name such as `muted`;
- give each quantitative claim an information relationship (a chart, axis,
  connector, direct label, or spatial comparison), not a decorative list of
  numbers;
- vary the dominant composition across a six-or-more-page deck and record an
  intentional exception when repetition serves the story;
- keep one clear reading anchor per page; do not fill unused space with equal-
  weight cards or captions.

These are authoring heuristics, not claims of aesthetic correctness. The
deterministic review may report bounded warnings; the Agent must inspect and
repair them or record why they are intentional.

## Editorial minimal (`editorial-minimal`)

- Build pages around one claim and one carefully chosen piece of evidence.
- Use whitespace and typographic contrast to establish sequence.
- Vary between statement, annotated evidence, and short comparison pages.
- Avoid decorative card grids, repeated centered slogans, and ornamental data.

## Enterprise data review (`enterprise-data-review`)

- Lead with the operating question and the variance that changes action.
- Keep chart scales comparable and label material values directly.
- Alternate overview, driver, exception, and decision pages.
- Avoid dashboard walls, equal emphasis across all metrics, and tiny legends.

## Technical architecture (`technical-architecture`)

- Show boundaries, ownership, flow direction, state, and failure recovery.
- Introduce complexity in stages; keep symbols and connector semantics stable.
- Pair topology with a concrete decision or operational consequence.
- Avoid decorative clouds, unexplained arrows, and diagrams without a reading path.

## Visual narrative (`visual-narrative`)

- Use image sequence, crop, scale, and pacing to carry the story.
- Give every image a factual or emotional job and preserve readable focal points.
- Use text as orientation, interpretation, or evidence—not a second full story.
- Avoid stock-image filler, repeated hero layouts, and captions that restate titles.

## Academic research (`academic-research`)

- Separate question, method, evidence, uncertainty, interpretation, and contribution.
- Make citation and figure provenance visible without overwhelming the claim.
- Use consistent notation and comparable analytical views.
- Avoid pretending exploratory results are causal or final.

## Brand launch (`brand-launch`)

- Establish a recognizable motif, reveal the product promise, then prove it.
- Use product imagery and demonstrations as evidence rather than decoration.
- Control tempo through contrast between reveal, detail, proof, and invitation.
- Avoid slogan stacks, feature inventories, and visual novelty without meaning.

## Combining mechanisms

Combine mechanisms only when both have clear jobs. Examples:

- technical architecture + enterprise data review for an operational platform decision;
- academic research + visual narrative for a public research talk;
- brand launch + editorial minimal for a restrained product announcement.

Resolve conflicts inside `designGrammar`. The grammar must name the final
palette roles, typography roles, spacing, density rhythm, motif, imagery,
chart treatment, and invariants for this deck.
