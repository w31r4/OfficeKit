## 1. PPJ contract

- [x] 1.1 Add `shadow.skewY` and `textShadowSkewY` to the PPJ schema and capability registry, then parse the JSON contracts successfully
- [x] 1.2 Extend the OpenSpec/native capability notes with the strict 1/60000-degree and `ky` bounds

## 2. Codec and source-bound lifecycle

- [x] 2.1 Parse, project, and author direct rich-text `skewY` while keeping other shadow transforms fail-closed; verify codec build succeeds
- [x] 2.2 Accept `textShadowSkewY` in the edit plan and splice only `outerShdw/@ky`; verify stale, mixed-transform, malformed, and unknown graphs stay opaque

## 3. Evidence and synchronization

- [x] 3.1 Add the focused authored/source-bound regression for round trip, changed-part scope, byte preservation, and second projection
- [x] 3.2 Update generated capability/Skill references and backlog/coverage notes, then run OpenSpec validation and the narrow affected tests
