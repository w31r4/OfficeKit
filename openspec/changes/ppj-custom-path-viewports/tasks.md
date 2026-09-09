## 1. Viewport lifecycle

- [x] 1.1 Add schema/lowering/projection/source support; verify mixed/default/large extents, canonical projection, source add/change/remove/no-op, unchanged commands/flags/other geometry/text/frame/non-target ZIP members and invalid/mask rejection in focused native tests.

## 2. Documentation and checks

- [x] 2.1 Update Help, registry, shapes reference, backlog and coverage; regenerate docs and pass preview diagnostics, portability/reference sync and strict OpenSpec checks.

Verification (2026-09-10): final SDK 8.0.128 CustomGeometry/Connector/PpjPreviewAuthoredSceneTests passed 128/128, zero skipped. Maximum width and height both reproject; a different native leaf kind still rejects the same oversized numeric value. Initial fixture numeric-node handling was corrected, and the nativeLeaf 1e9 schema mismatch was fixed for width/height only. Preview input assessment, capability coverage, presentation Skill maintenance, matrix freshness (57 boundaries/278 leaves), portability (255 files), reference sync (333 files), strict OpenSpec and diff checks passed. Concurrent internal painter files were preserved. No wire change, NativeAOT rebuild or host appearance claim.
