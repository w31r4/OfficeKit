# PPJ chart text style

## Why

PPJ currently exposes only `fontSize` for chart titles and axis tick labels.
This makes the language look broad while forcing authored charts back to host
defaults for font identity, emphasis, and color. It also discards a bounded
canonical text profile that OfficeKit can preserve and edit safely.

## What Changes

- Extend the existing chart text style with Latin and East Asian font family,
  bold, italic, RGB/theme-resolved color, and direct opacity.
- Compile and import one exact DrawingML `a:rPr` / `a:defRPr` profile.
- Issue one source-bound chart-text-style capability and lower only the title
  and axis style fields it declares.
- Regenerate the PPJ manual and update focused chart guidance.

## Impact

- Additive wire-v2 fields only; no protocol-version change.
- Shared chart wire state remains usable by XLSX and PPTX.
- Rich or irregular native text graphs remain source-owned and fail closed.

