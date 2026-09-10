## Context

See proposal.md. PptxParagraphLayoutCodec owns left margin and hanging indentation, using raw integer classification, a 51206400 EMU bound and explicit removal choices. PresentationTextParagraph currently uses field numbers through 34. The source compiler separates paragraph fields with masks, exact capabilities and per-paragraph patches.

## Goals / Non-Goals

**Goals:** Add the missing physical right paragraph margin using the same units, precision and source preservation as the left side.

**Non-Goals:** Text-box inset changes, logical start/end margins, inherited margin resolution or automatic reflow.

## Decisions

- Use rightIndent for the public point value and a:pPr/@marR for the native physical right margin. Direction does not swap the left and right fields.
- Extend the wire with oneof right_margin: margin_right_emu = 35 and no_margin_right = 36. This matches left_margin semantics: unset preserves native source state, a true removal choice clears the modeled direct attribute. Existing field numbers/version remain unchanged.
- Reuse the native 0..51206400 EMU margin bound and nearest-even point rounding. Validate authored point values before rounding; equivalent projected values round back to the exact source EMU.
- Classify namespace-free marR raw integer tokens before SDK numeric access. Invalid/out-of-range tokens stay residual, are omitted from semantic state and reject replacement. Compare existing modeled values before assigning to preserve plus signs and padding.
- Add rightIndent to the existing paragraph precedence, projection, exact authority and independent mutation mask. Only changed paragraphs receive removal intent.
- Normalize the new removal choice out of semantic hashes, as for left margin. Keep the layout reader/writer shared with list and table text.
- Preview diagnoses the unhandled right-margin native field; native export and re-import provide the field experiment, while host line wrapping remains partial.

Native mapping: [Microsoft Open XML SDK RightMargin](https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.drawing.textparagraphpropertiestype.rightmargin?view=openxml-3.0.1).

## Risks / Trade-offs

- Zero may collapse into absence → inspect attribute presence through zero, removal and restoration.
- Shared layout changes may disturb left/hanging properties or unknown XML → compare non-target XML/ZIP and run existing shared paragraph regressions.
- Published binaries lack the new field until the normal build → distinguish source/protobuf evidence from NativeAOT release and host layout acceptance.

The native removal experiment reproduced an equal-semantic-hash bypass for an unknown source marR. Keep the removal choice normalized for postwrite comparison, but require native writer dispatch and changed-element validation for the explicit right-margin removal intent, including nested groups and unchanged-source clone refusal.
