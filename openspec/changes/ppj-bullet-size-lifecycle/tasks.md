## 1. Size contract and source editing

- [x] 1.1 Add mutually exclusive point/relative/follow-text schema, lowering, projection and exact source authority; verify authored bounds/precision and text/shape add/set/switch/remove/restore with fresh native projections.
- [x] 1.2 Preserve malformed or ambiguous native size declarations and equivalent numeric spelling; verify no-op bytes, unrelated edits, rejected replacements and non-target XML/ZIP preservation.

## 2. Discovery and minimal verification

- [x] 2.1 Update Help, registry, generated manual/matrix, text guidance and F-03 backlog; verify generation checks and direct-point versus relative/follow-text preview diagnostics.
- [x] 2.2 Run focused bullet/list/table/master regressions, Skill checks and strict OpenSpec validation; record actual results and remaining host limitations, and verify the publication diff contains only this change.

## Evidence

The four focused managed tests first reproduced unsupported source edits,
missing follow-text schema and malformed-number projection failure. After the
implementation, fractional source edits exposed a native precision mismatch;
normalizing the requested wire value fixed it. Final focused result: 4/4 pass.

Related bullet/paragraph/list/table/master regression selection: 36 passed,
2 stale numbered-marker rejection assertions failed because the combined
formatting and direct-style removal is now supported. They now check the
actual candidate's fresh projection and byte-identical no-op; isolated rerun
passes 2/2. The fixture helper's fixed old numbering scheme was corrected in
that new assertion. No production behavior was relaxed for the old tests.

Preview input/coverage, both generated-document checks, Skill portability
(255 files), reference sync (333 files) and strict OpenSpec validation pass.
Direct point-size preview stays partial for layout; ratio/follow-text sizing
reports unavailable. No NativeAOT rebuild, full npm test or host acceptance.
