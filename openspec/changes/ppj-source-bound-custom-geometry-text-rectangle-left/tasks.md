## 1. Native leaf contract

- [x] 1.1 Add the bounded `customGeometryTextRectangleLeftEmu` leaf to PPJ native-leaf projection and the capability registry, issuing it only for recognized custom geometries with canonical in-frame literal rectangle edges.
- [x] 1.2 Extend native-leaf normalization and edit-plan proofing so rectangle order, direct/private representation, source token, and shape-local bound are revalidated before any write; reference-backed and invalid targets remain fail-closed.

## 2. Source-bound write path

- [x] 2.1 Add the SlidePart XML token splice and readback for `customGeometryTextRectangleLeftEmu`, preserving direct numeric versus private-guide representation while changing only the left coordinate.
- [x] 2.2 Add a focused authored/imported/source-bound/reprojection regression for one literal left edge, representation preservation, stale binding, and invalid input.

## 3. Evidence and gates

- [x] 3.1 Update F-04 backlog, coverage, PPJ schema, and generated reference evidence, then run strict OpenSpec validation, the focused codec test, JSON/review/reference sync smoke checks, and `git diff --check`.
