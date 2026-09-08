## 1. Contract

- [x] 1.1 Add `text.fontFamilyEastAsia` to the bounded formal target/reference and verify the existing owner vocabulary remains valid.
- [x] 1.2 Extend semantic target validation while preserving source-bound fail-closed boundaries; verify invalid-source diagnostics still fire.

## 2. Compiler and review evidence

- [x] 2.1 Resolve East Asian font values through the declared precedence chain and preserve the legacy Latin fallback; verify authored output contains the expected native typeface.
- [x] 2.2 Extend the focused formal round-trip/review fixture with theme-hit and default-fallback evidence; verify projected rich text restores both values.

## 3. Bookkeeping and gates

- [x] 3.1 Update coverage and K-05/F-03/F-15 backlog wording to record the bounded East Asian owner and retain broader font fallback residuals; verify reference sync.
- [x] 3.2 Run strict OpenSpec validation, focused C# and review tests, schema/reference/protocol gates, and `git diff --check`; record the results before closing the change.
