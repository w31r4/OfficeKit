## 1. Arrow lifecycle

- [x] 1.1 Add the capability vocabulary and source compiler/projector path; verify one focused fixture covers add/change/none/omission, preserved bindings/sizes, source no-op and SlidePart-only reprojection, plus denied capability.

## 2. Discoverability

- [x] 2.1 Update schema descriptions, capability/Help, focused shapes reference and F-04/coverage evidence; run related connector tests, generated docs, affected preview checks and strict OpenSpec before atomic publication.

Evidence (2026-09-10, isolated base 94e3006f): native Connector, PpjPreviewAuthoredSceneTests and FreePositionedLines filter passes 67/67, zero skipped. One original-source fixture separately adds startArrow, changes endArrow, changes to open, sets none and omits it. Native retained dimensions, removed-end absence, exact frame anchors, byte-exact no-op and SlidePart-only fresh projection are asserted. The existing native-site fixture changes to open and compares its original StartConnection XML; endpoint-edit authority remains denied. Authored open compiles to native arrow and fresh projection restores open. Tampered capability is rejected by nativeRef.stale before candidate emission.

Initial test corrections: scope the outline lookup to the connector because the target shape also has an outline; assert the actual earlier stale-nativeRef diagnostic for a modified capability list. Neither correction relaxes the successful edit or no-candidate checks. Discovery also found the shared open/arrow enum mismatch, fixed in authored/source lowering and projection.

Maintainer, generated matrix, preview scene SVG, capability coverage, Skill portability (255 files), root reference sync (333 files), whitespace and strict OpenSpec checks pass. No wire change or proto regeneration. Source-library evidence only; no NativeAOT rebuild or host/full-repository acceptance. Full F-04 and P0/P1 backlog remain open.
