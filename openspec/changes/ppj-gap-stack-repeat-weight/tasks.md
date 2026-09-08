## 1. Contract

- [x] 1.1 Add bounded optional `repeat.layout.weights` to the PPJ schema and reference, then verify the schema parses and the documented direction/count semantics are present.
- [x] 1.2 Carry weights through repeat parsing and reject incompatible direction, anchor, count, finiteness, and positivity cases with stable paths; verify validator diagnostics cover the invalid cases.

## 2. Authored expansion and evidence

- [x] 2.1 Lower valid horizontal and vertical weighted repeats into deterministic gap-separated slots while preserving omitted-weight behavior; verify the focused authored compile/project test passes.
- [x] 2.2 Add a minimal `[1,2,1]` authored round-trip and an invalid-weight assertion; verify projected ordinary frames have the expected ratio and gap.

## 3. Bookkeeping and gates

- [x] 3.1 Update coverage and the K-08/F-16 backlog to record the bounded weighted-stack owner and keep solver/source-bound residuals explicit; verify the references remain synchronized.
- [x] 3.2 Run strict OpenSpec validation, focused C# tests, schema/reference/review/protocol gates, and `git diff --check`; record the results before closing the change.
