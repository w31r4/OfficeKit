## 1. Contract and projection

- [x] 1.1 Add `textDefaultShadowScaleY` to the PPJ schema, native leaf registry,
      capability reference, and generated manual; verify JSON and maintainer
      contract checks.
- [x] 1.2 Project an existing paragraph default outer-shadow `sy` as a
      1/100000 scale leaf and advertise the source-bound style capability;
      verify the leaf value and paragraph binding in a focused fixture.

## 2. Source-bound codec

- [x] 2.1 Extend edit-plan validation and expected-value proof for the signed
      scale token; verify out-of-range, stale, and unsupported graphs fail
      closed.
- [x] 2.2 Splice only `a:outerShdw/@sy` in the owning SlidePart; verify sibling
      effects, text topology, and non-target package parts remain unchanged.

## 3. Evidence and publication

- [x] 3.1 Add the authored/imported source-bound lifecycle regression covering
      precision, changed-part scope, Open XML validity, re-projection, and
      rejection cases; verify the focused test passes.
- [x] 3.2 Update the F-03 backlog and coverage evidence, validate the OpenSpec
      change strictly, run the narrow codec/portability/reference gates, then
      make one atomic commit and push it to `origin/main`.
