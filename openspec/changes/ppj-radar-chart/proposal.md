# PPJ standard radar chart

## Why

PPJ already declares `radar` as a chart type, but the authored compiler rejects
it and the native chart codec cannot project it. The language therefore
advertises a visual primitive that cannot be built, inspected or continued.

## What Changes

- Add a wire-v2 radar chart enum value.
- Compile PPJ `chartType: "radar"` to a native editable standard radar plot.
- Project canonical native radar plots back to PPJ.
- Reuse the existing category/value, series stroke, marker, label, axis and
  source-bound chart-edit contracts.
- Keep filled radar, custom radar styles and waterfall charts fail closed.

## Impact

- Additive protobuf enum; protocol version remains 2.
- The PPJ schema is unchanged because `radar` is already declared.
- The shared ChartSpace codec owns one bounded standard-radar representation,
  while only the PPJ/PPTX route enables it in this change. The Spreadsheet
  public surface remains unchanged.
