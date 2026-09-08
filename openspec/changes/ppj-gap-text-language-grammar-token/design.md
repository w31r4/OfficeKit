# Design

## Contract

- `textStyle.language` accepts either a literal `languageTag` or a
  `grammarTokenRef`.
- A token must declare kind `string`; its resolved value must be a valid
  bounded BCP-47 language tag.
- Run-level formal precedence continues to select the winning source before
  this value resolution.
- The existing `PresentationTextRun.language` and `PresentationTextStyle.language`
  wire fields remain unchanged.

## Boundaries

The token does not select a font, infer a language from text, alter writing
direction, or make imported theme/master language owners editable. Invalid
token kinds and invalid resolved tags fail closed.
