## Why

Direct rich-text runs can carry the DrawingML outer-shadow vertical skew (`ky`), but PPJ currently drops that value and cannot edit it source-bound. This leaves a remaining P0 F-03 text-effect field after the bounded scale and horizontal-skew increments.

## What Changes

- Expose `run.style.shadow.skewY` in authored and projected PPJ.
- Add the `textShadowSkewY` native leaf with canonical signed `ky` precision and strict degree bounds.
- Allow a source-bound edit to splice only `a:outerShdw/@ky` in the owning SlidePart and reproject the changed value.
- Keep mixed transforms, extra effects, malformed values, and unknown descendants opaque or fail closed.
- Add a focused authored/source-bound regression and update capability, coverage, backlog, and presentation Skill references.

## Capabilities

### New Capabilities

- `ppj-source-bound-text-shadow-skew-y`: Direct rich-text run vertical shadow skew with bounded authoring, projection, and source-preserving editing.

### Modified Capabilities

None.

## Impact

The PPJ schema and capability registry, NativeAOT presentation parser/projector/compiler/edit-plan codec, focused OfficeKit codec tests, generated presentation capability matrix, presentation references, coverage, and the PPJ gap backlog are affected. No wire version change or host-PowerPoint rendering claim is required.
