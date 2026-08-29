# Shapes and diagrams

Decide the page's core message before drawing. A shape earns space only when it
encodes data, a relationship, a boundary, focus, or identity.

## Choose geometry by meaning

- Use position, scale, alignment, and shared baselines for comparison.
- Use lines and connectors for direction, dependency, sequence, or causality.
- Use bounded regions for real states, phases, ownership, or physical areas.
- Use native vectors for a deliberate identity motif when they carry the same
  meaning across the deck.
- Use a chart or table instead of a diagram when quantity or exact values are
  the relationship.

PPJ shape elements use a stable `id`, explicit `frame`, typed `geometry`, and
an optional named style. Connector endpoints should bind stable element IDs
when the relationship belongs to those objects. Keep arrow direction and label
placement unambiguous.

Custom geometry is justified only when a preset cannot express a necessary
semantic form. Keep its path finite and editable. Do not trace a decorative
illustration into arbitrary geometry when an image or SVG asset is the honest
carrier.

## Protect reading and evidence

Order `pages[].elements[]` from back to front. Keep evidence-bearing lines,
markers, labels, numbers, axes, intervals, and sources above fills or clear of
them. A foreground shape may overlap a background field; two evidence objects
must remain traceable.

When a collision occurs, repair the composition: adjust the frame, anchor the
label, reduce an honest fill's opacity, use a local mask behind text, or change
the carrier. Do not falsify scale or separate truly related series merely to
silence an overlap check.

## Strictly forbidden

- card walls or equal rounded panels used as default hierarchy;
- colored side-strip cards, pills, badges, and button-like labels as filler;
- random circles, rings, arrows, blobs, or nodes added to make a page "rich";
- decorative process diagrams with no process relationship;
- connectors that cross labels, values, or unrelated objects;
- large empty containers whose border does all the organizing;
- a universal `box`, `card`, or `metricPanel` component driving every page.

User-provided card-based brand systems and imported layouts may be preserved.
New shapes inside them still need a declared role and clear reading order.
