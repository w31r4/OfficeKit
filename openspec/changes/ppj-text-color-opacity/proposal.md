# PPJ Text Color Opacity

## Why

PPJ already accepts alpha-bearing colors, while the authored Presentation
compiler rejects alpha on text runs and default text styles. This forces an
Agent to replace an ordinary typographic hierarchy with heavier shapes or fully
opaque text. The native DrawingML codec can model one direct `a:alpha` child
without exposing arbitrary effects, so the public language should own that
bounded state.

## What Changes

- Preserve optional text-color opacity on Presentation run and default-run
  style wire messages.
- Compile PPJ alpha-bearing text colors into direct DrawingML solid paint.
- Project recognized direct RGB or theme text alpha back into PPJ.
- Keep transformed or multi-effect text paint source-owned and fail closed.
- Document opacity as a hierarchy tool rather than a substitute for readable
  contrast.

## Impact

This is an additive wire-v2 change and a bounded PPJ compiler expansion. It
does not add a new DSL field because alpha is already part of the strict PPJ
color value. It does not issue a new source-edit capability.
