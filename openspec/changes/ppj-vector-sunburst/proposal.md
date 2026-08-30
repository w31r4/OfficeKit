# PPJ vector sunburst

## Why

PPJ now authors hierarchical treemaps but cannot express the same bounded
forest as concentric part-to-whole rings. PPTD exposes a sunburst keyword yet
black-box output is a PNG with private source metadata. OfficeKit can compile
the semantic hierarchy into editable native custom-geometry sectors instead.

## What changes

- Add `sunburst` to the PPJ chart vocabulary.
- Reuse the aligned `parents` hierarchy channel and its finite invariants.
- Add bounded radial style and deterministic native sector lowering.
- Preserve exact embedded-PPJ recovery and truthful custom-shape group fallback.
- Add one focused example and Agent-facing selection guidance.

## What does not change

- No ChartPart, raster fallback, raw OOXML, arbitrary expression or wire-version
  change is introduced.
- Arbitrary imported arcs are not inferred to be a semantic sunburst.
