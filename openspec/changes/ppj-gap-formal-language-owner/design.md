# Design

## Contract

- `text.language` is an additional target supported by the bounded authored
  `stylePrecedence` profile.
- The value remains the existing PPJ BCP-47 `languageTag` string; this slice
  does not add grammar-token references or a new wire field.
- The compiler uses the same declared first-hit source order as the existing
  text scalars and writes the winning value to direct `a:rPr/@lang`.
- If no `text.language` rule is declared, the historical inline/default/style
  lookup order remains unchanged.

## Boundaries

The slice does not infer language from Unicode, choose a font, change text
direction, perform shaping/proofing, or write inherited/source-bound theme or
master owners. Unsupported language tags continue to fail schema validation.
