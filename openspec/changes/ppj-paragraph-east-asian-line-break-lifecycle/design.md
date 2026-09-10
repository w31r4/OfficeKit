## Context

See proposal.md for F-03 scope. PptxParagraphPropertiesCodec already handles
optional boolean paragraph settings with raw token preservation. The SDK maps
[EastAsianLineBreak to eaLnBrk](https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.drawing.textparagraphpropertiestype.eastasianlinebreak?view=openxml-3.0.1).
Office [uses this setting for East Asian line-start/end rules](https://learn.microsoft.com/en-us/openspecs/office_standards/ms-oi29500/9b34280e-538e-4811-8af9-761d34f88f20),
with separate kinsoku settings governing word breaks. A boolean field must retain
that setting without claiming to compute its layout or materializing a default.

## Goals / Non-Goals

**Goals:** Complete the direct boolean lifecycle through the shared paragraph owner.

**Non-Goals:** Inherited resolution, custom kinsoku lists, font measurement and
host line layout. Delivery comprises codec source and generated bindings.

## Decisions

- Add optional wire bool `east_asian_line_break = 40`, preserving existing numbers.
  Absence clears modeled direct source state and preserves an unknown token, as
  for other direct paragraph booleans. A separate removal flag is unnecessary.
- Read raw 0/1/false/true before SDK BooleanValue access. Preserve unchanged
  native spelling and reject replacement of unrecognized tokens.
- Extend exact capability validation, all independent paragraph diff masks and
  per-paragraph mutation; reuse shared list/table/master paths.
- Verify authored precedence, boolean presence, text/shape source lifecycles with
  Latin word breaking, hanging punctuation, wrap and break inlines kept independent,
  plus unknown-token and explicit input/native preview diagnostics.

## Risks / Trade-offs

- False or absence collapsed to a default → explicit false and absent-state re-import experiments.
- Shared writer or masks affect adjacent settings → source XML/ZIP and related paragraph/list/master/table regressions.
- A serialized flag mistaken for kinsoku evaluation → explicit partial preview diagnostics.
