# PPJ chart-marker fill opacity

## Why

PPJ marker fills use the alpha-aware color type, while the authored compiler
rejects alpha and the native marker projection stores only RGB. This makes the
declared language broader than its executable chart surface.

## What Changes

- Add presence-aware opacity to the bounded native marker fill.
- Read and write one direct DrawingML `a:alpha` child.
- Compile and project alpha-bearing PPJ marker fills.

## Impact

- Additive Office wire-v2 field only.
- The shared native chart marker codec recognizes the same bounded profile for
  Presentation and Spreadsheet charts; no new JavaScript Spreadsheet API is
  introduced.
