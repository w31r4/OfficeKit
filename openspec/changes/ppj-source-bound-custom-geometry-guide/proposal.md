## Why

F-04 still keeps an imported custom geometry's calculated guide graph source-owned. That is correct for formulas and topology, but it leaves a safe, common case—an existing `a:gd fmla="val N"` guide—without an independently editable PPJ field. A bounded literal-guide leaf closes one concrete geometry gap without pretending to edit the surrounding guide graph.

## What Changes

- Expose each existing, direct custom-geometry guide whose formula is exactly `val N` as a `customGeometryGuide` native leaf.
- Allow source-bound PPJ to change only that guide's numeric value by token-splicing its `fmla` attribute in the owning SlidePart.
- Preserve the guide name, formula topology, paths, handles, connection sites, text rectangle, and all non-target XML; reject missing, ambiguous, calculated, or out-of-range guides.
- Add a focused authored/imported/source-bound/reprojection regression and update the F-04 coverage record.

## Capabilities

### New Capabilities

- `ppj-source-bound-geometry-guide`: A bounded native leaf for editing an existing literal custom-geometry guide value while retaining the rest of the DrawingML geometry graph.

### Modified Capabilities

<!-- No existing OpenSpec capability files exist in this checkout. -->

## Impact

- PPJ native-leaf projection, edit-plan validation/proof, and SlidePart XML token splicing.
- `src/ppj/capability-registry.json` and the PPJ gap/coverage documentation.
- Native codec focused tests; no protocol version change is required because native leaves are already represented by the existing edit-plan wire shape.
