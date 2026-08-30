## Why

The capability registry already maps `presentation.slideSize` to
`$.design.canvas`, and the native PPTX codec already performs a bounded
canvas-only source edit. Source-bound PPJ currently projects width and height
but gives them no nativeRef, does not lower changes, and does not reject them.
An Agent can therefore build a modified PPJ successfully while receiving the
unchanged PPTX bytes. This is a false public contract.

## What Changes

- Add an optional canvas nativeRef and bounded `setCanvas` capability.
- Project `setCanvas` from every exact imported PPTX canvas.
- Lower capable width/height changes to the existing source-bound `p:sldSz`
  writer using point-to-EMU conversion.
- Preserve every slide/layout/master/object coordinate; no implicit scale or
  reflow occurs.
- Extend the existing comprehensive PPJ transaction and generated guidance.

## Capabilities

### New Capabilities

- `ppj-canvas-parity`: Truthful source-bound `$.design.canvas` mutation.

### Modified Capabilities

None. The PPJ schema ID and Office wire protocol version remain unchanged.

## Impact

PPJ schema/model validation, projection, source-bound lowering, generated Skill
reference, review guidance, coverage, and one existing test are affected. The
existing native PPTX canvas writer is reused unchanged.
