## Context

The shared native chart model already contains `SpreadsheetChartTextStyleArtifact` and `SpreadsheetChartLineOptionsArtifact`. XLSX ChartSpace import/build/patch owns title font size plus canonical line-chart `smooth`, `grouping`, and `varyColors`, but `PptxChartCodec` currently drops title/line state while converting between `PresentationChart` and `SpreadsheetChartArtifact`. PPJ consequently cannot declare, compile, or recover those fields.

## Goals / Non-Goals

**Goals:**

- Expose title font size and line-chart smoothing/color variation as strict PPJ state.
- Reuse the existing shared native chart messages and ChartSpace codecs.
- Preserve explicit smooth `false` separately from omission and preserve direct `varyColors: true`.
- Restore recognized state during PPTX-to-PPJ projection.
- Keep imported style mutations capability-bound and fail closed.

**Non-Goals:**

- New chart families, legend typography, per-label typography, or raw ChartSpace fields.
- General font family/color/bold state where the native owner does not exist.
- Automatic style mutation of arbitrary source-bound charts.

## Decisions

1. Add `titleTextStyle`, `smooth`, and `varyColors` to `chartStyle` so inline and named chart styles share one inheritance path. A second chart-format object would duplicate existing PPJ style semantics.
2. Map `titleTextStyle.fontSize` to `PresentationChart.TitleTextStyle`; map line behavior to `PresentationChart.LineOptions`. The bridge copies those messages to and from the shared worksheet-chart model rather than reimplementing XML.
3. Permit `smooth` and `varyColors` only for line charts. Other chart types reject the fields before native output changes.
4. Project `smooth` whenever native presence exists, including explicit false. Project `varyColors` only when true because false and omission have the same canonical native representation in the bounded profile.
5. Leave source-bound chart style immutable in this tranche. The importer may expose the state for inspection and exact recovery, but a changed style without a dedicated native capability is rejected.

## Risks / Trade-offs

- [Named style precedence becomes ambiguous] → Resolve every new field through the existing inline-over-named `FirstProperty` path.
- [Presentation bridge silently drops native state] → Copy both messages in both conversion directions and cover build/reimport in the existing integrated PPJ test.
- [A source edit appears supported merely because it projects] → Keep `ApplyChartElement` and issued capabilities unchanged; document the authored/reimported versus source-bound boundary.
- [Broad testing cost grows] → Add assertions to the existing integrated PPJ test only; run one filtered test plus C# build and Skill/OpenSpec consistency checks.
