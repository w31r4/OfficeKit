## 1. Field lifecycle

- [x] 1.1 Add setConnectorType through schema/validator/projector/compiler; verify type transitions with actual native geometry, retained endpoints/arrows/bindings, source no-op, SlidePart-only fresh projection and invalid/denied requests.

## 2. Discoverability

- [x] 2.1 Update schema/Help/capability and focused shapes reference plus backlog/coverage, run focused connector and affected JS/generated checks, and validate OpenSpec before atomic publication.

Evidence (2026-09-10, isolated base ddc4213f): Connector and PpjPreviewAuthoredSceneTests filter passes 67/67, zero skipped. The new focused case switches straight to elbow to curved to straight, reprojects each exact candidate before the next edit, checks native preset identity and retained endpoints/anchor/arrow/stroke, proves source no-op at each stage, and compares all ZIP members (only target SlidePart changes). Invalid enum and tampered capability return no candidate. Existing native-site/arrow/coordinate regressions remain green.

Maintainer, generated matrix, preview scene SVG, preview capability coverage, portability (255 files), root reference sync (333 files), whitespace and strict OpenSpec pass. No wire field or native writer changed. These are source-library results, not a new NativeAOT build, host routing test or full repository acceptance. Full F-04 and P0/P1 backlog remain unfinished.
