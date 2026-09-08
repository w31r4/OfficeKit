## 1. Native leaf contract

- [x] 1.1 Add the bounded `customGeometryAdjustmentHandlePolarYEmu` leaf to PPJ native-leaf projection and the capability registry, issuing it only for direct ordered polar handles with canonical literal in-frame y coordinates.
- [x] 1.2 Extend native-leaf normalization and edit-plan proofing so the handle index/kind, direct attributes, position child, source token, and shape-local bound are revalidated before any write; reference-backed and invalid targets remain fail-closed.

## 2. Source-bound write path

- [x] 2.1 Add the SlidePart XML token splice and readback for `customGeometryAdjustmentHandlePolarYEmu`, replacing only the selected `a:ahPolar/a:pos/@y` value while preserving guide identity, range bounds, x, handles, and package topology.
- [x] 2.2 Add a focused authored/imported/source-bound/reprojection regression for one literal polar position y, an in-frame replacement, identity preservation, stale binding, and out-of-frame input.

## 3. Evidence and gates

- [x] 3.1 Update F-04 backlog, coverage, PPJ schema, and generated reference evidence, then run strict OpenSpec validation, the focused codec test, JSON/review/reference sync smoke checks, and `git diff --check`.
