## Context

`PptxCodec.ApplySourceBoundSlideSize` already compares the requested dimensions
with the exact imported `p:sldSz`, writes positive signed 32-bit EMUs, clears a
stale preset type, and changes only `ppt/presentation.xml`. PPJ's authored
compiler consumes the same canvas state, but the source-bound compiler stops
before this mature writer.

## Goals / Non-Goals

**Goals:**

- Make imported canvas state capability-bound and editable through PPJ.
- Preserve exact point semantics at the JSON boundary and deterministic EMU
  rounding at compile time.
- Report every page as affected because the presentation canvas changes while
  the page coordinate systems remain untouched.
- Recover the requested dimensions after a second projection.

**Non-Goals:**

- Scaling, reflowing, cropping, or repositioning page content.
- Changing master/layout/page coordinate values.
- Guessing an Office preset aspect-ratio name.
- Adding a separate command or imperative operation list.

## Decisions

### 1. Canvas owns its nativeRef

`design.canvas` gains optional `nativeRef`. Imported projection issues
`setCanvas` for `canvas.width` and `canvas.height`; source-free programs omit
the proof. The source SHA, revision, semantic object hash, and capability-set
hash bind the operation to the exact PPTX.

### 2. The existing field remains the operation

The Agent edits width and/or height in points. The compiler requires the
unchanged nativeRef and converts the values to EMUs with the same deterministic
rounding used by authored PPJ. There is no `resizeCanvas()` action object.

### 3. Layout consequence is explicit

Changing the canvas never rescales content. Review must render every page and
check for newly exposed margins, clipping, or changed composition. This is a
deliberately sharp primitive, not an automatic layout feature.

## Risks / Trade-offs

- [Silent no-op remains] -> Compare the fresh projected canvas and require
  `setCanvas` before mutating the native artifact.
- [Precision drift] -> Round point values once to signed 32-bit EMUs and recover
  points from those exact integers.
- [Agent assumes reflow] -> State the canvas-only boundary in `ppj.md`, review
  guidance, capability summary, and coverage.

## Migration Plan

Additive optional nativeRef plus one closed capability operation and two fields.
Existing authored and source-bound PPJ remains schema-valid.

## Open Questions

None.
