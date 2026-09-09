## Context

`OpenXmlChartTrendlineLabelCodec` and `XlsxChartTextStyleCodec` live in the shared codec assembly; `PptxTextCodec` depends on slide-owned relationships in the presentation assembly. Reuse the shared chart style parser/writer, not the slide codec or a new presentation authoring engine.

## Goals / Non-Goals

Provide a coherent chart text field for literal paragraphs/runs/breaks with bounded direct chart styles. Formula/field/hyperlink ownership, body/list formatting and effect graphs remain source-owned. This does not close all rich-text or F-07 gaps.

## Decisions

- Extend `label.text` from string/token to also accept `{paragraphs:[{style?,runs:[{text,style?}|{break:true,style?}],endStyle?}]}`. Reuse `chartTextStyle`; alignment is paragraph-only. Text runs accept string tokens. Empty paragraphs and empty text runs are valid. Use typed breaks instead of control characters.
- Add a label rich-text wire field plus three appended messages for body/paragraph/run. Keep the old literal field. Canonicalize one unstyled nonempty run of at most 255 characters into the legacy literal state in both compilation and reading, so equivalent native text has one semantic representation. All other structured state remains rich text.
- Preserve paragraph/run order, whitespace, explicit false character overrides and styles. Empty native bodyPr/lstStyle and empty optional pPr/rPr/endParaRPr carry no modeled style and normalize to absence. Nonempty unknown properties, hyperlinks, fields and extensions fail closed; native no-op retains original bytes.
- Share mechanical PPJ traversal with caller-provided string/style conversion delegates. Reuse existing chart style grammar token resolution in authored/source-bound paths and the existing style projector. String token resolvers gain an explicit `allowEmpty` option used only for rich runs; existing callers retain their nonempty constraint.
- Bound resources to 4096 paragraphs, 16384 inlines and 1,048,576 UTF-16 characters per label. Validate counts, oneof content and character styles before writing. These are explicit codec budgets, not renderer acceptance.
- Use existing `setChartSeriesAnalytics` for transitions among structured/literal/automatic text, preserving other label/chart/package state.

## Risks / Trade-offs

- Reusing a style parser can silently accept unmodeled text/comments → preflight the entire rich tree and strictly check child order/attributes before exposing editability.
- Paragraph alignment on a run would be dropped by native XML → reject it in schema and wire validation.
- Equivalent simple structured text reprojects as string → document and test canonicalization, including no-op comparison.
