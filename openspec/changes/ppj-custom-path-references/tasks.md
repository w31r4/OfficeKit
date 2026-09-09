## 1. Referenced path lifecycle

- [x] 1.1 Add custom-only parameter unions, reference-aware lowering/projection and source gate; verify all command slots, source literal/reference replacement and dependency evaluation, invalid references/arcs, line/mask boundaries and unchanged non-target data with focused native geometry/connector tests.

## 2. Discoverability and checks

- [x] 2.1 Update Help, registry, shapes reference, backlog and coverage; regenerate docs and pass preview diagnostics/scene, portability/reference-sync and strict OpenSpec validation.

Verification (2026-09-10): final SDK 8.0.128 CustomGeometry/Connector/PpjPreviewAuthoredSceneTests passed 127/127, zero skipped. The line boundary has paired numeric-success/reference-rejection inputs; its initial control fixture had a malformed stroke color and was corrected before the final run. Preview input assessment, scene SVG, capability coverage, presentation Skill maintenance, matrix freshness (56 boundaries/278 leaves), portability (255 files), reference sync (333 files), strict OpenSpec and diff checks passed. No wire change, NativeAOT rebuild or host geometry claim.
