## Context

Chart text styles are shared by native XLSX/PPTX ChartML and PPJ. Their exact style parser currently permits sz/b/i/u but rejects lang. PPJ already defines languageTagOrToken and PptxLanguageTag contains a pure bounded tag validator with no slide dependency.

## Goals / Non-Goals

Complete language state on chartTextStyle and all existing consumers of that style. Other character fields such as altLang, complex effects, dictionaries and host proofreading remain separate work.

## Decisions

- Add optional string `language=12` to SpreadsheetChartTextStyleArtifact and reference existing languageTagOrToken in chartTextStyle. Preserve casing and optional presence, including explicit en-US; omission removes the native attribute. The existing 2–63-character lexical profile remains the contract, without installed-locale lookup.
- Compile the existing pure PptxLanguageTag source in the shared codec and exclude its duplicate compilation from the presentation project. Keep the same validator and error behavior for existing presentation callers; chart validation uses IsValid and a chart-specific error.
- Include language in meaningful-style detection, parsing, writing and semantic comparison, and reject it in the special global-font-family-only profile. A language-only chart style is meaningful. Paragraph/run/end styles automatically reuse these helpers.
- Update authored style construction, per-field grammar precedence, source-bound conversion and projection. Pass chart language defaults to vector title runs when no more-specific run language exists, and resolve language in the shared vector label builder for other chartTextStyle consumers.
- Use native chart and rich-label lifecycle fixtures plus a vector-title native-language assertion. Keep JS imports conservative and verify unrelated JS chart edits preserve wire language fields.

## Risks / Trade-offs

- Clearing en-US as a generic text default would lose explicit chart state → retain HasLanguage in chart semantics and native projection.
- Missing language in a style-presence check would drop language-only owners → test them and the global-font-family guard directly.
- Concurrent preview compiler edits → build and publish a snapshot containing only this field's hunks.
