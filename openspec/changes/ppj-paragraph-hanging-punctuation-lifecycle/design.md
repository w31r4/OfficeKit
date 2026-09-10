## Context

See proposal.md for F-03 scope. Shared PptxParagraphPropertiesCodec already handles
direct optional boolean direction, including raw source token preservation.
The SDK exposes DrawingML `hangingPunct` as the misleadingly named boolean
[TextParagraphPropertiesType.Height](https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.drawing.textparagraphpropertiestype.height?view=openxml-3.0.1).

## Goals / Non-Goals

**Goals:** Use the existing paragraph owner and edit rules for a complete direct
boolean lifecycle, with public terminology matching hanging punctuation.

**Non-Goals:** Inherited setting resolution, CJK line-break rules, font measurement
or host punctuation placement. Source and generated bindings are the delivery surface.

## Decisions

- Add optional wire bool `hanging_punctuation = 38`, preserving existing numbers.
  Absence clears a modeled direct source attribute and preserves an unknown one,
  matching direction; a second removal flag is unnecessary for this contract.
- Classify raw 0/1/false/true before SDK boolean access and retain lexical spelling
  if the value is unchanged. Use Height only inside the native codec with a comment
  explaining its mapping; PPJ and wire use hangingPunctuation terminology.
- Extend exact capability validation, all independent paragraph diff masks and
  per-paragraph mutation. Reuse shared list/table/master codecs rather than adding
  another writer. Preserve first-line hanging indent independently.
- Add small authored, text/shape lifecycle and unknown-token experiments plus explicit
  input/native preview diagnostics. Keep generated references synchronized.

## Risks / Trade-offs

- False mistaken for absence → authored precedence, wrapper deletion and native re-import checks.
- Shared writer or mask changes affecting neighbors → source XML/ZIP and related paragraph/list/master/table regressions.
- Serialized setting mistaken for layout support → explicit partial preview diagnostics and no host fidelity claim.
