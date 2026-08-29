# Shared visual floor

This is the single cross-scenario authority for choosing visual form. It is not
a house style. The selected scenario, design authority, and deck-specific
Design Grammar decide palette, typography, theme, density, imagery, and page
silhouettes. User templates, brand rules, and preserved source design take
precedence.

## Communication and evidence

- Give every page one audience task and one primary conclusion. A page may use
  several facts only when they jointly prove that conclusion.
- Make the reading order and information hierarchy evident at a glance.
- Let charts, images, geometry, text, and negative space serve the current
  argument. Decoration is not a substitute for evidence.
- Never invent facts, data, quotes, cases, people, outcomes, or sources.
- Match the medium: `live` favors staged comprehension, `reader` preserves
  self-contained evidence, and `hybrid` must work both spoken and unattended.
- Keep titles minimal and specific. Make executive summaries scan through
  hierarchy and light dividers. A closing page resolves the opening and adds no
  new evidence.

## Choose form from meaning

For every newly composed page, decide in this order:

```text
audience task → claim and evidence → dominant relationship
              → visual carrier → geometry roles → composition
```

- Name the dominant relationship as quantity/trend, sequence, hierarchy,
  causality, comparison, spatial context, object evidence, or explicitly
  `none`. Do not begin from a favorite shape, helper, or layout.
- Choose the carrier that makes that relationship easiest to understand: chart,
  image, table, diagram, typography, native vector, or a deliberate mixture.
- Map quantitative change to charts, exact lookup to tables, sequence/hierarchy/
  causality to diagrams, physical identity or context to sourced imagery, and a
  purely verbal proposition to typography. Use native geometry to express a
  simple boundary, relation, annotation, or identity—not as generic enrichment.
- Give each added geometric element a role: encode data, connect a relationship,
  annotate evidence, establish a boundary or page field, create focal hierarchy,
  or maintain identity and continuity. Ease of drawing is not a role.
- A structural diagram should express one primary relationship. Every connector
  must communicate direction, dependency, sequence, or causality.
- Do not add circles, rings, arrows, nodes, blocks, panels, or cards merely to
  fill space or make a page feel richer. If removing an element only makes the
  page emptier, remove it.
- Preserve authoritative template and imported geometry. The role test applies
  to new additions and redesigns; it does not rewrite untouched source design.

## Composition and fit

- Make titles, body copy, labels, and sources readable in the rendered deck.
  Shorten or recompose before shrinking type.
- Keep all content inside the slide with no unintended overlap, clipping, or
  wrapping. A one-line title must remain one line.
- Make negative space intentional: it should focus attention, separate ideas,
  or control pacing. A large unused region with no such role is a composition
  defect, not sophistication.
- Give each evidence page a dominant carrier. Let a chart, image, diagram,
  table, typographic statement, or native-vector composition own the field.
- For data, choose a chart whose form matches the relationship, direct-label
  important values when practical, and use color to direct attention rather
  than decorate every series.
- Keep evidence-bearing lines, markers, labels, axes, connectors, arrowheads,
  and source-critical image regions visible. Treat data marks as protected
  foreground: no band, surface, image, shape, annotation, or other series may
  hide a line segment, marker, label, axis, interval, or error bar the audience
  must read.
- When truthful series overlap, preserve the relationship and resolve label or
  fill conflicts by re-anchoring labels, reserving marker clearance, choosing
  honest axis ranges, using bounded transparency or local masks, putting the
  line above pale bars, or choosing a valid alternate chart form. Do not split
  a shared plot merely to make an overlap warning disappear.
- Separate aligned panels only when overlaying different units would mislead or
  no truthful, legible shared encoding survives the target-host render. Label
  every independent scale explicitly.
- Use icons only when they encode structure or meaning, and keep one coherent
  icon language within a deck.
- Preserve native editability and render every changed page before delivery.

## Containers are not a default composition

On an ordinary OfficeKit self-directed page without an authoritative template
or brand requirement, do not use:

- repeated rectangles or rounded rectangles as a card wall;
- cards to perform ordinary alignment, hierarchy, or whitespace;
- colored-side-strip cards or equal-width, equal-height panel grids;
- pills, badges, tabs, or button-like text boxes as decoration;
- a generic `box()`, `card()`, or `metricPanel()` as the page's main language;
- `roundRect + outline + shadow` as a default component.

Express parallel content through the relationship and carrier chosen above,
using scale, position, baseline, light rules, numbering, connectors, charts,
tables, images, and negative space.

The prohibition does not apply when the treatment is authoritative or
structural: a user-supplied card template or brand system, preserved imported
cards, a real product UI screenshot, table cells, a chart plot area, an image
crop, or an explicitly requested dashboard. A page-scale field or frame still
needs a declared boundary, focus, or continuity role; its size is not an
exemption.

## Final judgment

Do not prescribe a universal color, font, corner radius, shadow, image count,
information density, or page outline here. Record those choices in the current
Design Grammar and judge them against the communication task. Structured checks
may locate risks; only an understood render or human review can support an
aesthetic conclusion.
