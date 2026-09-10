## 1. Right indent field and lifecycle

- [x] 1.1 Add schema, right-margin wire choice, authored precedence and native projection with bounded EMU precision; verify zero, range/rounding and precedence in focused native experiments and regenerate bindings.
- [x] 1.2 Implement exact source authority, per-paragraph add/set/remove/restore and native numeric preservation/refusal and explicit native removal dispatch; verify original-source text/shape lifecycle and non-target XML/ZIP checks.

## 2. Discovery and integration

- [x] 2.1 Update Help, registry, generated manual/matrix, focused text guidance and F-03 backlog; verify generation and explicit right-margin preview diagnostics.
- [x] 2.2 Run proto:check, affected layout/direction/tab/list/table/master regressions, Skill and strict OpenSpec checks; record results and verify an isolated atomic publication diff.

## Verification evidence

- Focused right-indent experiments: 6/6 passed across ordinary text/shape lifecycle, zero and range/rounding, authored precedence, source spelling and invalid native coordinates. Native removal checks cover modeled, absent and unknown attributes for both top-level and grouped text.
- The native removal experiment first reproduced a no-op bypass (unknown right margin was accepted); writer dispatch and postwrite change classification now honor the explicit removal intent and the experiment passes.
- npm run proto:check, preview input/wire/coverage, generated manual/matrix, Skill maintenance, portability (255 files), reference sync (333 files) and strict OpenSpec checks passed.
- Native source was compiled and exercised with SDK 8.0.128. No full npm test, NativeAOT rebuild or PowerPoint host reflow acceptance was performed; inherited margins and host line wrapping remain partial.
- Related native regressions: 25/25 passed. The initial run passed 24; the list-style fixture was updated because right margin is now modeled, and its rerun passed 1/1. It proves right-margin deletion while retaining a real a:extLst payload. Other checks cover left/hanging indent, direction, tab stops, fields/breaks, list/master defaults, table text, native groups and unchanged group cloning.
- Isolated staged snapshot: preview input/wire/coverage, generated manual/matrix, Skill maintenance and strict OpenSpec checks passed with only this field included. Native source/test files match the tested working files.
