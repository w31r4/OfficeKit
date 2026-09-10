## 1. Field and source lifecycle

- [x] 1.1 Add schema/wire/native mapping and authored precedence; regenerate bindings and verify signed bounds, EMU precision, zero/absence, independent tab lists/literal tabs and invalid input in minimal native experiments.
- [x] 1.2 Add exact authority and per-paragraph edits; verify text/shape source add/set/remove/restore, single-field style removal, unknown-token refusal and non-target XML/ZIP preservation.

## 2. Discovery and integration

- [x] 2.1 Synchronize Help, registry, generated manual/matrix, focused text guidance and F-03 backlog; verify explicit input/native preview diagnostics for zero and nonzero values.
- [x] 2.2 Run proto:check, affected tab/paragraph/list/master/table regressions, Skill and strict OpenSpec checks; record results and verify the scoped atomic publication candidate.

## Validation evidence

- Focused native experiments: 4/4 passed using SDK 8.0.128 and Open XML SDK 3.5.1. Signed endpoints, nearest-EMU precision, zero/absence, authored precedence, text/shape source lifecycles and unknown-token preservation/refusal passed.
- Explicit tab-list deletion preserves defaultTabSize; field edits preserve literal tabs, explicit tab lists, numeric spelling, neighboring paragraphs and non-target XML/ZIP. All 42 paragraph diff masks remain independent.
- Initial run was 3/4 because the fixture expected an empty tab list for noTabStops; the existing contract removes the list. Corrected that assertion and added the independent list-deletion preservation check; production behavior was unchanged by this fixture correction.
- Preview input/native diagnostics for negative, zero and positive values, scene wire and generated Skill maintenance checks passed. Measured tab placement and host layout remain partial; no NativeAOT rebuild is claimed.

- Related tab-stop/right-indent/paragraph/list/master/table regressions: 19/19 passed. proto:check, Skill portability/reference-source sync, preview capability coverage and strict OpenSpec validation passed.
