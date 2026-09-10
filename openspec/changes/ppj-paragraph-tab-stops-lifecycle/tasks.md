## 1. Field lifecycle and source preservation

- [x] 1.1 Complete bounded ordered positions, choice precedence, exact authority and per-paragraph add/set/clear/remove/restore; verify authored bounds and fresh source text/shape lifecycle tests.
- [x] 1.2 Classify native tab lists safely, preserve equivalent spelling and retain unmodeled lists during independent edits; verify malformed-source refusals and non-target XML/ZIP preservation.

## 2. Discovery and verification

- [x] 2.1 Update schema, Help, registry, generated manual/matrix, text guidance and F-03 backlog; verify the generation checks and explicit tab-layout preview diagnostic.
- [x] 2.2 Run the focused field/break/tab and shared paragraph/list/table/master regressions, Skill and strict OpenSpec checks; record actual results and verify the scoped publication diff.

## Verification evidence

- New native lifecycle experiments: 7/7 passed across ordinary text/shape, empty native lists, grouped clear dispatch, precision/bounds, precedence and 13 unmodeled source profiles. The first focused run also passed the existing PPJ tab-stop author/reprojection case.
- Related native regressions: 10/10 passed across fields/breaks/tabs, rich paragraphs, list defaults, master text, mixed-run and picture-bullet table cells, style precedence and native groups. Two stale expectations were corrected and rerun (2/2): exact bullet alpha is represented as an RGB object, and an opaque group retains its proven direct-frame editing capability. The group test now verifies movement and re-import instead of only the flag.
- Preview input/unmapped-tab diagnostic and capability-coverage tests passed. Skill maintenance, generated matrix, portability (255 files), reference sync (333 files) and strict OpenSpec checks passed.
- Native source tests used SDK 8.0.128. No wire change, full npm test, NativeAOT rebuild or PowerPoint host layout acceptance was performed. Inherited tab placement and host typography remain partial.
- The isolated staged snapshot passed preview input assessment, capability coverage, generated manual/matrix and strict OpenSpec checks; only the tab-stop registry entry is included, preserving concurrent preview work outside this commit.
