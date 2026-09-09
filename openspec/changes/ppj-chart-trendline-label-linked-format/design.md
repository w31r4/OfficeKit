## Context

The shared trendline-label reader currently requires exactly formatCode plus an explicit false sourceLinked attribute; the writer always emits false. Number-format strings already support grammar tokens and source-bound analytics. See proposal.md for motivation.

## Goals / Non-Goals

Preserve the complete native attribute state on this label owner and allow its lifecycle. Axis/data-label formatting, workbook format synchronization and host rendering remain separate work.

## Decisions

- PPJ `numberFormatSourceLinked` is boolean or null and requires `numberFormat`. True writes `1`, false or omitted PPJ property writes `0`, null omits the native attribute. Projection canonicalizes false by omitting the PPJ property; true and native omission project true/null. Null avoids guessing a default when importing an absent native attribute.
- Use an additive wire enum `SpreadsheetChartNumberFormatLink` with Unspecified=0 (legacy explicit false), Source=1, Omitted=2, and label field `number_format_link=7`. A zero default retains the established behavior of old wire clients; a nullable boolean alone would change their meaning. Unknown enum values and nondefault link state without a format fail validation.
- Both compiler paths resolve the field through one small label helper. Existing typed label semantics and `setChartSeriesAnalytics` carry changes. No new native edit operation or workbook dependency traversal is needed.
- Native XML accepts 0/false/1/true and absence, preserving absence as state. Invalid attributes/content still keep the owner opaque. Label-only edits retain non-target package bytes.

References: [Number format attribute](https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.drawing.charts.numberingformat) describes the standard default; [Office implementation note](https://learn.microsoft.com/en-us/openspecs/office_standards/ms-oe376/baa94e34-1d3e-4d59-a139-d9955fcdcf79) documents a historical default difference and states that Office does not use this attribute on trendline labels. This change preserves the native state and does not promise visual inheritance.

## Risks / Trade-offs

- Mistaking the flag for an active workbook synchronization feature → make its owner-specific behavior explicit in field documentation.
- Collapsing native absence into a guessed default → represent it with null and test absence after unrelated label changes.
- Concurrent preview protocol work → validate and publish only this field from an isolated snapshot.
