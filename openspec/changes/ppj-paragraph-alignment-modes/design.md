## Context

The shared native paragraph codec and PPJ projection map five values. `PptxTableCodec.BoundedTableParagraphProperties` also has a five-value allowlist. Schema paragraph styles are shared by direct text, defaults and structured text consumers. Preview already reports unsupported alignment painting for modes other than left/center/right.

[Microsoft's TextAlignmentTypeValues definition](https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.drawing.textalignmenttypevalues?view=openxml-3.0.1) identifies JustifiedLow as justLow and ThaiDistributed as thaiDist.

## Goals / Non-Goals

**Goals:** Complete the seven-value paragraph enum through the existing compiler, projection and source lifecycle.

**Non-Goals:** New wire fields, native leaf editing syntax, chart-wide typography changes, NativeAOT release and host glyph shaping.

## Decisions

- Use `justifyLow` and `thaiDistributed` in PPJ, extending the existing readable `justify`/`distributed` naming. Raw OOXML spellings remain internal.
- Extend the shared raw-token reader, native writer and projection normalization together. Reuse existing exact field authority and absence-based deletion.
- Extend the table allowlist so reading a newly modeled mode cannot make an otherwise supported table opaque. Keep every other table constraint intact.
- Reuse the alignment lifecycle fixture with seven values and replace its now-modeled negative tokens with invalid ones. Add bounded table and default-owner round trips; verify preview through existing partial diagnostics instead of approximating Arabic or Thai justification.

## Risks / Trade-offs

- Shared owners can regress during projection → check bounded table and master/list defaults in addition to direct text/shape lifecycle.
- Readers of older PPJ schemas reject the new enum values → the change is additive; old values and wire layout remain unchanged.
- Structural fidelity does not establish glyph placement → document partial preview and retain separate host verification.
