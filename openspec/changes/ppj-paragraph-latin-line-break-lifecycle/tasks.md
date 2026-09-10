## 1. Field and source lifecycle

- [x] 1.1 Add schema/wire/native mapping and authored precedence; regenerate bindings and verify true/false/absence, independent wrap/break settings and invalid input in native experiments.
- [x] 1.2 Add exact source authority and per-paragraph lifecycle; verify text/shape removal/restoration, style-wrapper removal, boolean spelling, unknown-token refusal and non-target XML/ZIP preservation.

## 2. Discovery and integration

- [x] 2.1 Update Help, registry, generated manual/matrix, focused text guidance and F-03 backlog; verify explicit preview input/native diagnostics and documentation checks.
- [x] 2.2 Run proto:check, affected paragraph/list/master/table regressions, Skill and strict OpenSpec checks; record results and verify the scoped atomic publication.

## Verification evidence

- Native Latin-line-break experiments: 4/4 passed (text/shape lifecycle with existing break inlines, boolean presence and precedence, independent wrap, unknown-source preservation/refusal). SDK 8.0.128; snapshot-free export/re-import, original-source deletion/restoration and target XML/ZIP checks are included.
- Preview input/native scene diagnostics passed for true and false. All 40 paragraph diff masks were checked for independent field comparisons.
- proto:check, preview input/wire/coverage, generated manual/matrix, Skill portability (255 files), reference sync (333 files) and strict OpenSpec checks passed.
- Source and generated bindings only: no full npm test, NativeAOT rebuild or PowerPoint host word-breaking acceptance was performed. Inherited defaults and measured word/line layout remain partial.
- Related native regressions: 17/17 passed, covering hanging punctuation, direction, font alignment, fields/breaks/tab stops, list defaults, master text styles and table mixed-run/picture-bullet source preservation.
- Isolated staged snapshot: preview input/wire/coverage, generated manual/matrix, Skill maintenance and strict OpenSpec checks passed with only this field included. Native source/test files match the tested working files. The publication includes 22 scoped files; concurrent preview/source-deletion work is retained separately.
