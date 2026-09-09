## Why

F-07 still preserves multi-run trendline labels as source-owned. The existing label text only accepts one unstyled literal run, so agents cannot express or edit styled labels with paragraphs and line breaks.

## What Changes

- Extend `label.text` to accept ordered `paragraphs` with text/break runs, paragraph defaults, run styles and end-paragraph styles.
- Reuse chart text-style fields and grammar tokens; retain string/token compatibility and the existing automatic-text behavior when text is omitted.
- Implement authored/source-bound replacement, deletion, recreation and fresh projection for ordinary and categorical combo labels; reject unsupported text graphs without flattening.
- Add a focused native lifecycle experiment and synchronize the chart reference/coverage.

## Capabilities

### New Capabilities
- `ppj-chart-trendline-rich-text`: Native styled trendline label paragraphs and runs.

### Modified Capabilities

None.

## Impact

PPJ schema, artifact protocol/bindings, shared chart text codec/style helpers, both PPJ compiler routes and projector, focused tests and presentation documentation. No slide-codec dependency or new edit operation. Formula references, hyperlinks, fields, complex body/list styles and effects remain separate F-07 gaps.
