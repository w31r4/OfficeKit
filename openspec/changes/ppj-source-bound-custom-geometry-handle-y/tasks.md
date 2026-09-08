## 1. Native leaf contract

- [x] 1.1 Add the bounded `customGeometryAdjustmentHandleYEmu` leaf to PPJ native-leaf projection and the capability registry, issuing it only for direct ordered XY handles with canonical in-frame literal y positions.
- [x] 1.2 Extend native-leaf normalization and edit-plan proofing so the handle index/kind, direct position structure, source token, and shape-frame range are revalidated before any write; formula-backed and invalid targets remain fail-closed.

## 2. Source-bound write path

- [x] 2.1 Add the SlidePart XML token splice and readback for `customGeometryAdjustmentHandleYEmu`, replacing only the selected `a:ahXY/a:pos/@y` value while preserving guide identity, bounds, x, other handles, and package topology.
- [x] 2.2 Add a focused authored/imported/source-bound/reprojection regression for one literal XY handle y position, handle identity, stale binding, and out-of-frame input.

## 3. Evidence and gates

- [x] 3.1 Update F-04 backlog, coverage, PPJ schema, and generated reference evidence, then run strict OpenSpec validation, the focused codec test, JSON/review/reference sync smoke checks, and `git diff --check`.
