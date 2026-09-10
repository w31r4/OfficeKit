## Context

See proposal.md. Native inner-shadow fields already preserve optional blur/distance/direction/alpha, and chart text has optional PPJ geometry. Paragraph default construction/projection currently uses the ordinary required-geometry profile. The glow increment supplies isolated direct-effect readers and source-local node patching.

## Goals / Non-Goals

Preserve native optional state within paragraph defaults, with the existing effect value ranges and exact source authority. Other text-style properties keep their constraints, and run/shape/image inner-shadow authoring retains its current required geometry. Host rendering and other effect lifecycles remain separate.

## Decisions

Create a paragraph-default text-style schema view whose property references reuse the existing textStyle constraints; only innerShadow uses the existing optional-geometry color profile. Keeping the ordinary textStyle contract avoids silently expanding run/shape/image authoring. Preserve the color/gradient conflict rule. The native schema resolver and generated manual must resolve these property pointers inside existing definitions; neither may treat a field name as a standalone definition.

Reuse optional native construction and omit missing geometry during paragraph-default projection. Preserve omitted versus explicit opacity, RGBA, theme identity and source-bound grammar precedence. Resolve transformed colors to RGB; round angles before wrapping into 0..360 degrees.

Generalize the proven direct glow patch helper only enough to share it with innerShadow. Isolated readers validate the target node; replacement retains its position, insertion uses native ordering, and deletion removes only that node and an empty attribute-free list. Unsupported direct inner-shadow graphs remain source-owned.

Compare optional inner-shadow presence explicitly because generated protobuf equality equates missing numeric values and zero. Skip unchanged direct runs using serialized equality, which retains presence and avoids rewriting their source effects during paragraph-only edits. The existing textDefaultInnerShadowDistanceEmu leaf accepts up to 1,270,000,000 EMU in the schema so the 100,000pt endpoint can reproject without broadening other leaf ranges.

## Risks / Trade-offs

Schema aliases lose diagnostics or constraints → retain property references and exercise existing default-style cases plus explicit inner-shadow preview diagnostics.

Missing geometry becomes zero during compile/project → test imported omission, each optional leaf removal, explicit zero and fresh projection.

Mixed effects or theme precision change during an unrelated edit → preserve source XML/ZIP and neighboring paragraph/direct-run fixtures.

Known whole-default-style baseline failure → retain its existing exclusion and report the exact related filter, without promoting full F-03 or host acceptance.
