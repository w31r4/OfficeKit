## 1. PPJ and codec contract

- [x] 1.1 Add `textShadowRotateWithShape` to the native-leaf schema, registry, generated matrix, and reference; verify JSON parsing and the matrix check pass
- [x] 1.2 Extend the strict direct-run shadow safety and projection paths to admit `rotWithShape` as the sole transform; verify the focused projection test passes

## 2. Source-bound lifecycle

- [x] 2.1 Add native-leaf whitelist, boolean validation, proof/readback, dispatch, and `rotWithShape` token patch support; verify the source-bound edit test changes only the owning SlidePart
- [x] 2.2 Add authored, explicit-false, unsupported-topology, stale, XML, ZIP-byte, and second-projection regression coverage; verify the focused test filter passes

## 3. Evidence and gates

- [x] 3.1 Update the presentation text Skill, coverage, and P0/P1 backlog evidence; verify the presentation maintainer check passes
- [x] 3.2 Run OpenSpec strict validation, codec build, focused tests, schema/matrix checks, portability/reference-sync, and preview capability/field-limit gates; record any environment-only skip (the broader reference-skills render suite was not run because this increment uses the focused gates)
