## 1. Native leaf contract

- [x] 1.1 Add the bounded `customGeometryConnectionSiteYEmu` leaf to PPJ native-leaf projection and the capability registry, issuing it only for direct ordered sites with canonical in-frame literal y tokens.
- [x] 1.2 Extend native-leaf normalization and edit-plan proofing so the site index, direct position structure, source token, and shape-frame range are revalidated before any write; formula-backed and invalid targets remain fail-closed.

## 2. Source-bound write path

- [x] 2.1 Add the SlidePart XML token splice and readback for `customGeometryConnectionSiteYEmu`, replacing only the selected `a:pos/@y` value while preserving x, angle, formulas, paths, handles, site order, and package topology.
- [x] 2.2 Add a focused authored/imported/source-bound/reprojection regression for one literal y coordinate, one formula-backed sibling, stale binding, and out-of-frame input.

## 3. Evidence and gates

- [x] 3.1 Update F-04 backlog, coverage, PPJ schema, and generated reference evidence, then run strict OpenSpec validation, the focused codec test, JSON/review/reference sync smoke checks, and `git diff --check`.
