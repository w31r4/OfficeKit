## Context

See proposal.md. OpenXmlChartTrendlineCodec currently rejects every trendlineLbl child. Existing point-label codecs already own a one-run literal text profile, numeric formats and direct paint/line/text styles; PPJ compilers have corresponding grammar resolution.

## Goals / Non-Goals

Own the label container and its existing bounded appearance/text vocabulary. Automatic equation calculation, manual layout, rich text/formula graphs and visual fidelity remain separate F-07 work.

## Decisions

- Add a dedicated wire label message and one optional trendline message field. Append the new message so existing descriptor indexes stay stable. Using a general point-label override would expose irrelevant position/visibility fields.
- Reuse point-label text/shape-property helpers and existing text-style codecs. Require native child order, unique children, literal text and sourceLinked=false; accept an absent or empty layout without authoring a manual layout.
- Project and resolve all five label properties through existing PPJ chart grammar. Object presence distinguishes an automatic empty label from no label.
- Include label semantics in trendline comparison. Existing analytics list replacement supplies lifecycle and preserves unchanged native trendlines.
- Preserve the original wire label when the JS workbook adapter rebuilds an imported trendline for an unrelated chart edit; that adapter retains its existing public syntax.
- Verify one ordinary and one combo lifecycle plus unmodeled-owner rejection and adjacent trendline/error-bar regressions. Check target native nodes and ZIP preservation rather than adding a rendering acceptance process.

## Risks / Trade-offs

- Unknown label content could be erased → recognize only the complete bounded label profile before enabling analytics.
- Default label presence could disappear → retain an empty wire message and PPJ object explicitly.
- Reused point-label syntax could leak unrelated options → use a dedicated schema/message and copy only the five owned properties.
- Shared XLSX parser behavior changes → include its trendline codec regression; do not claim a new JavaScript workbook authoring surface.

## Migration Plan

The wire addition is backward compatible within the existing protocol. Regenerate bindings, update discoverability/docs, and record managed round-trip evidence. Full F-07 and host/visual validation remain open.
