# PPJ chart gradient paint

## Why

PPJ already uses one fill union for authored shapes, backgrounds, tables and
chart surfaces, but chart series still collapse to an RGB scalar. This makes a
common visual primitive appear missing even though OfficeKit can now recognize
the same bounded DrawingML gradient on imported presentation objects.

## What Changes

- Extend the shared chart fill message with the existing bounded gradient.
- Let chart areas, plot areas and series use `none`, `solid` or `gradient` PPJ
  fills, including stop opacity.
- Project canonical imported fills and allow capability-bound source edits.
- Preserve legacy solid-series fields for internal compatibility.
- Keep theme, pattern, image, parameterized and effect-bearing fills source
  owned and fail closed.

## Impact

- Additive wire-v2 fields; no protocol-version change.
- Ordinary and combo PPTX charts share the implementation.
- No new PPJ syntax is invented: the existing `fill` union becomes truthful at
  every declared chart location.
