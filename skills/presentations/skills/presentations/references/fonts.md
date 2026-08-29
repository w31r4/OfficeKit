# Presentation typography guidance

Typography is a role system, not a single approved font. Select a family from
evidence available on the target host and record substitutions that affect
metrics or readability. A user or brand design system wins over these defaults.

## Assign roles before choosing families

| Role | Job | Typical treatment |
| --- | --- | --- |
| Display | Cover or section promise | Highest contrast in size/weight; short line length. |
| Title | State the page conclusion | Sentence or claim, not a topic label; stable across the deck. |
| Body | Explain evidence and decisions | Comfortable line height, restrained weight, readable measure. |
| Data | Make values and units comparable | Tabular numerals when available; align decimal or unit columns. |
| Label | Identify axes, legends and diagram nodes | Smaller than body but never sacrificed below the delivery floor. |
| Source | Preserve provenance and time | Quiet contrast, still readable in the intended after-use. |
| UI/chrome | Page numbers, controls and metadata | Consistent and subordinate; do not compete with the claim. |

## Language and host evidence

- For CJK text, choose a family with verified glyph coverage and pair it with a
  Latin family only when the x-height, stroke contrast and numeral widths are
  visually compatible.
- Do not infer that a font name is installed because it is common elsewhere.
  Inspect the host/package evidence or choose a known fallback and record it.
- Avoid mixing more than two primary families in one deck unless a supplied
  design system requires it. A CJK fallback is an implementation detail, not a
  new visual role.
- Check punctuation, full-width characters, line breaking, numerals and mixed
  Latin/CJK labels in an actual render. Font substitution can change wrapping,
  baseline, chart labels and placeholder fit.
- Use a font-size floor appropriate to delivery. Shorten, split or recompose
  before shrinking a title, body or source below that floor.

## Rhythm and hierarchy

Use a small, explicit scale rather than many arbitrary sizes. A page should
make its reading order obvious through title/body/data contrast, alignment,
weight and spacing—not by placing every sentence in a different container.

Titles should carry the conclusion. Body copy should earn its line count. Data
labels should sit next to the mark they explain. Sources should remain attached
to the evidence they qualify. Keep line length, paragraph spacing and baseline
alignment consistent within a page archetype, then vary density intentionally
across the deck.

## Review before delivery

Review a representative title, body, data label, source line, and mixed CJK /
Latin page after export and re-import. Look for:

- substituted glyphs, missing symbols or fallback boxes;
- changed line breaks, clipped text, collapsed spacing or baseline drift;
- labels colliding with marks, lines, bars or image edges;
- source text becoming unreadable in the after-use medium;
- a hierarchy that depends on color alone or weight alone.

If a renderer substitutes a font, keep the fact in the task evidence and either
choose a verified compatible family, adjust the copy/geometry, or mark the
visual review as requiring human confirmation. Do not claim a font-specific
visual result that was not rendered.

