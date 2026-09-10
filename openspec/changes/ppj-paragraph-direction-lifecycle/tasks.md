## 1. Direction field and native lifecycle

- [x] 1.1 Add schema, optional wire presence and authored/projected paragraph direction; verify LTR/RTL, invalid values and style precedence with focused native tests and regenerated bindings.
- [x] 1.2 Implement exact source authority, per-paragraph add/set/remove/restore and raw RTL preservation/refusal; verify original-source text/shape lifecycle, unknown-token and non-target XML/ZIP checks.

## 2. Discovery and integration

- [x] 2.1 Update Help, registry, generated manual/matrix, focused text guidance and F-03 backlog; verify generation and explicit direction preview diagnostics.
- [x] 2.2 Run protocol generation consistency, affected paragraph/list/table/master regressions, Skill and strict OpenSpec checks; record results and verify an isolated atomic publication diff.

## Verification evidence

- Direction lifecycle experiments: 4/4 passed. Ordinary text/shape fixtures check explicit LTR/RTL, original-source removal, restoration, direction-only wrapper deletion on empty text, exact authority, source boolean spelling, neighboring XML/ZIP preservation and unknown-token refusal. Authored checks retain independent alignment/column direction and verify style precedence plus invalid inputs.
- Related native regressions: 22/22 passed across paragraph level/alignment/tab stops, field/break text, rich paragraph preservation, list defaults, master text and mixed-run/picture-bullet table cells.
- npm run proto:check passed against the staged generated binding. Preview wire, input direction diagnostics, registry coverage, Skill maintenance, generated matrix, portability (255 files), reference sync (333 files) and strict OpenSpec checks passed.
- SDK 8.0.128 compiled and exercised the native source. No full npm test, NativeAOT binary rebuild or PowerPoint bidi/typographic host acceptance was performed. Inherited direction and host shaping remain partial.
- The isolated staged snapshot passed preview input/wire/coverage, generated manual/matrix and strict OpenSpec checks. Publication includes only the new direction registry entry; concurrent preview and source-deletion work stays outside this commit.
