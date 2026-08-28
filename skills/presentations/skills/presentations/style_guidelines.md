# Shared visual floor

This is a cross-scenario quality floor, not a house style. The selected scenario,
design authority, and deck-specific Design Grammar decide palette, typography,
theme, density, geometry, imagery, and page silhouettes. User templates and
brand rules take precedence.

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
  and source-critical image regions visible. When truthful series overlap,
  preserve the relationship and resolve label or fill conflicts with bounded
  transparency, offsets, local masks, or a valid alternate chart form.
- Use icons only when they encode structure or meaning, and keep one coherent
  icon language within a deck.
- Preserve native editability and render every changed page before delivery.

## Card-based composition is forbidden by default

On an ordinary OfficeKit self-directed page without an authoritative template
or brand requirement, do not use:

- repeated rectangles or rounded rectangles as a card wall;
- cards to perform ordinary alignment, hierarchy, or whitespace;
- colored-side-strip cards or equal-width, equal-height panel grids;
- pills, badges, tabs, or button-like text boxes as decoration;
- a generic `box()`, `card()`, or `metricPanel()` as the page's main language;
- `roundRect + outline + shadow` as a default component.

Express parallel content through scale, position, baseline, light rules,
numbering, connectors, charts, tables, images, and negative space.

The prohibition does not apply when the card treatment is authoritative or
structural: a user-supplied card template or brand system, preserved imported
cards, a real product UI screenshot, table cells, a chart plot area, an image
crop, one page-scale background field or frame, or an explicitly requested
dashboard reproduction.

## Final judgment

Do not prescribe a universal color, font, corner radius, shadow, image count,
information density, or page outline here. Record those choices in the current
Design Grammar and judge them against the communication task. Structured checks
may locate risks; only an understood render or human review can support an
aesthetic conclusion.
