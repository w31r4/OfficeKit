## 1. Contract and projection

- [x] 1.1 Add `textDefaultReflectionSkewX` to the PPJ schema, native leaf
      registry, capability reference, and generated manual; verify JSON and
      maintainer contract checks.
- [x] 1.2 Project an existing paragraph default reflection `kx` as a signed
      1/60000-degree leaf and verify the value and paragraph binding in the
      focused fixture.

## 2. Source-bound codec

- [x] 2.1 Extend edit-plan validation and expected-value proof for canonical
      signed skew tokens; verify stale, malformed, and unsupported reflection
      graphs fail closed.
- [x] 2.2 Splice only `a:reflection/@kx` in the owning SlidePart; verify `sx`,
      `sy`, `fadeDir`, sibling effects, text topology, and non-target package
      parts remain unchanged.

## 3. Evidence and publication

- [x] 3.1 Add the source-bound lifecycle regression covering precision,
      changed-part scope, Open XML validity, re-projection, and rejection cases;
      verify the focused test passes.
- [x] 3.2 Update the F-03 backlog and coverage evidence, validate the OpenSpec
      change strictly, run the narrow codec/portability/reference gates, then
      make one atomic commit and push it to `origin/main`.
