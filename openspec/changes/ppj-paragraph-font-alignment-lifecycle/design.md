## Context

See proposal.md for the F-03 gap. PptxParagraphPropertiesCodec already shares
alignment/direction handling across direct paragraphs, list defaults and master
text styles. PPJ currently has no font-alignment field in schema or wire.
Microsoft documents [the fontAlgn attribute](https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.drawing.textparagraphpropertiestype.fontalignment?view=openxml-3.0.1)
and [all five enum values](https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.drawing.textfontalignmentvalues?view=openxml-3.0.1).

## Goals / Non-Goals

**Goals:** Preserve direct enum presence and unknown native state using the same
paragraph owner and independent source edit conventions as alignment/direction.

**Non-Goals:** Host font measurement, inherited alignment resolution, line
wrapping or automatic paragraph restructuring. Preview remains explicitly partial.

## Decisions

- Add optional wire string `font_alignment = 37`; use public names and retain
  absence. Like paragraph alignment/direction, absence clears a modeled direct
  attribute while preserving an unrecognized source token. A separate removal
  flag would introduce a second absence policy without a need for this enum.
- Classify raw `fontAlgn` before SDK enum access. Map auto/t/ctr/base/b to
  auto/top/center/baseline/bottom, validate authored/native requests, and preserve
  unknown tokens during unrelated edits. Do not coerce unknown to auto.
- Extend existing paragraph diff masks and exact field authority. Mutate only
  paragraphs whose fontAlignment changes; leave sibling values and source XML intact.
- Reuse the shared paragraph codec for list/table/master surfaces. Verify related
  regressions and the text/shape lifecycle rather than adding a separate writer.
- Add small native authored/lifecycle/unknown-source experiments and input/native
  scene diagnostic assertions. Keep generated manual and registry synchronized.

## Risks / Trade-offs

- Shared paragraph writer changes can affect defaults → focused list/master/table regressions.
- Omitting a diff mask could widen authority or rewrite neighbors → exact-capability and XML/ZIP footprint experiments.
- An exported hint is not measured host layout → explicit preview diagnostic and source-only delivery evidence.
