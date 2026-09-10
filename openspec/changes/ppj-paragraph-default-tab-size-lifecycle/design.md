## Context

See proposal.md for the F-03 gap. Shared paragraph properties already own direct
optional attributes and raw-token preservation, including East Asian line break.
The project's Open XML SDK 3.5.1 defines
[DefaultTabSize as Int32Value without a narrower validator](https://raw.githubusercontent.com/dotnet/Open-XML-SDK/v3.5.1/data/schemas/schemas_openxmlformats_org_drawingml_2006_main.json).
[The SDK property maps to defTabSz](https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.drawing.textparagraphpropertiestype.defaulttabsize?view=openxml-3.0.1).
Existing projected point coordinates normally round to six decimals; that can
put the minimum signed coordinate outside an exact public range.

## Goals / Non-Goals

**Goals:** Preserve the full direct coordinate and its source lifecycle, including zero.

**Non-Goals:** Inherited resolution, measured tab layout and host rendering. Deliver
codec source and generated bindings without rebuilding NativeAOT for this increment.

## Decisions

- Use points bounded by -2147483648/12700 through 2147483647/12700, with nearest-EMU
  conversion and ties-to-even. Accept native negative coordinates too; the SDK
  does not constrain this field to nonnegative margins. Invalid or out-of-range
  authored values reject before conversion.
- Add optional int32 `default_tab_size_emu = 41`. Presence distinguishes zero
  and absence; absence clears modeled direct state while preserving unknown raw
  tokens. No explicit removal oneof is needed for this direct optional attribute.
- Project by direct division by 12700 rather than six-decimal rounding, so both
  signed endpoints remain schema-valid and recover the original integer EMU.
- Parse raw integers before SDK value access, preserve equivalent lexical forms,
  and refuse replacing unknown tokens. Reuse shared paragraph read/write paths,
  exact capability checks, independent masks and per-paragraph mutation.
- Keep explicit tab lists, noTabStops, indentation and literal tab runs independent;
  diagnose both zero and nonzero default sizes until preview implements tab placement.

## Risks / Trade-offs

- [Boundary precision drift] → endpoint, half-EMU and source no-op/reprojection experiments.
- [Zero confused with absence] → zero-only style removal and original-source restoration.
- [Unrelated tab list or neighbor changes] → compare non-target XML/ZIP and preserve raw spelling.
- [A stored coordinate mistaken for layout proof] → explicit preview partial diagnostics.
