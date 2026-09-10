## Context

See proposal.md for F-03 scope. Shared PptxParagraphPropertiesCodec already handles
optional boolean direction and hanging punctuation with raw token preservation.
The SDK maps [LatinLineBreak to latinLnBrk](https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.drawing.textparagraphpropertiestype.latinlinebreak?view=openxml-3.0.1).
Office's [documented default differs from the standard](https://learn.microsoft.com/en-us/openspecs/office_standards/ms-oi29500/9b34280e-538e-4811-8af9-761d34f88f20),
so preserving absence is necessary rather than materializing either default.

## Goals / Non-Goals

**Goals:** Complete the direct boolean lifecycle through the shared paragraph owner.

**Non-Goals:** Inherited value resolution, word-break algorithms, font measurement
or host line layout. Delivery comprises codec source and generated bindings.

## Decisions

- Add optional wire bool `latin_line_break = 39`, preserving existing numbers.
  Absence clears modeled direct source state and preserves an unknown token, as
  for direction/hanging punctuation. A separate removal flag adds no required state.
- Read raw 0/1/false/true before accessing SDK BooleanValue. Preserve original
  spelling for unchanged values and refuse to overwrite unrecognized tokens.
- Extend exact capability validation, all independent paragraph diff masks and
  per-paragraph mutation; retain existing shared list/table/master paths.
- Verify with authored precedence and presence, text/shape source lifecycles,
  unknown-token experiments and explicit input/native preview diagnostics. Include
  existing break inlines to prove that the setting does not rewrite text topology.

## Risks / Trade-offs

- False or absence collapsed to a default → explicit false and absent-state re-import checks.
- Shared writer or masks affect unrelated paragraph content → native XML/ZIP and adjacent paragraph/list/master/table regressions.
- A serialized setting mistaken for line layout → explicit partial preview diagnostics.
