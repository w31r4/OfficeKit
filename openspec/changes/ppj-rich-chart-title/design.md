## Context

PowerPoint stores a chart title inside the ChartPart as DrawingML paragraphs
and runs under `c:title/c:tx/c:rich`. OfficeKit already owns the same bounded
paragraph/run vocabulary for ordinary presentation text. The chart wire,
however, currently retains only flattened text plus one uniform style.

## Goals / Non-Goals

**Goals:**

- Honor the existing PPJ `textContent` contract for authored chart titles.
- Preserve stable chart identity, frame, series, axes, relationships, and
  non-title ChartPart XML while editing a rich title.
- Recover compiler-owned paragraphs and runs during PPTX projection.
- Keep string titles deterministic and backward compatible.

**Non-Goals:**

- Formula-backed chart titles, arbitrary chart text-layout properties, WordArt,
  title hyperlinks, or a generalized ChartPart XML editing API.
- Making an unsupported third-party chart editable merely because its title is
  readable.

## Decisions

### 1. Structured title is additive

`PresentationChart` receives `title_body`; the existing `title` remains the
flattened compatibility value and must equal the structured body. Older wire
payloads and PPJ string titles continue through the existing path.

### 2. Reuse the Presentation text vocabulary

The chart codec reuses `PresentationTextBody` and the existing paragraph/run
validation. It serializes that bounded state into the ChartPart rich-text
container rather than inventing a second typographic model.

### 3. Uniform title style becomes a default

For a structured title, chart `titleTextStyle` supplies defaults only where a
run has not declared its own font, emphasis, size, or color. The emitted title
is self-contained, so reimport can recover the same visible run semantics.

### 4. Imported edits are profile-bound

Projection exposes structured title content only when the native rich-text
container has one body-properties node, one list-style node, bounded paragraph
and run topology, and no external hyperlink relationship. Otherwise the title
remains flattened for inspection and the chart is not advertised as editable.

### 5. Replace only the owned rich-text subtree

An accepted source-bound title edit replaces `c:tx/c:rich` while preserving
the surrounding title layout and every non-title ChartPart node. It does not
rebuild the chart or its embedded workbook.

## Risks / Trade-offs

- [The shared text vocabulary contains features inappropriate for chart
  titles] -> Reject title hyperlinks and noncanonical title containers before
  compile or edit, while retaining the broadly useful typography fields.
- [Uniform title style and run style conflict] -> Treat the uniform style as a
  default and let explicit run properties win.
- [Third-party rich title carries unknown extensions] -> Do not issue an edit
  capability; preserve it through the source package.

## Migration Plan

Add the wire field, codec, compiler/projector support, guidance, and one focused
round-trip assertion. Existing PPJ and PPTX output without structured chart
titles remains unchanged. Rollback removes the additive field and restores the
explicit authored rejection.

## Open Questions

None. Formula-backed titles require a separate source-bound capability.
