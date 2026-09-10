## 1. Numbering format contract

- [x] 1.1 Enumerate supported schemes, authorize canonical/alias source edits and preserve native marker attributes; verify the authored catalog and text/shape source edit, restore, combined startAt, alias and negative cases with focused regressions.

## 2. Discoverability and shared behavior

- [x] 2.1 Synchronize Help, schema descriptions, registry, generated PPJ reference/matrix, text guidance and F-03 backlog; verify numbering preview diagnostics, affected paragraph/bullet/list tests, generated checks, portability/reference sync and strict OpenSpec validation.

Validation (2026-09-10): focused tests pass 3/3; one authored document covers all 41 schemes and five aliases, while text/shape fixtures verify source edits, canonical reprojection, exact native no-op for an equivalent alias, restoration, combined startAt changes, missing/invalid requests and non-target XML/ZIP preservation. The related startAt, paragraph, bullet/list/master and table-picture-bullet filter passes 22/22 with zero skips. Schema and native catalogs match all 41 values. Preview diagnostics, generated manual/matrix, portability/reference sync and strict OpenSpec pass. Native refs remain immutable: the positive fixture keeps the exact projected reference; authority removal is only a negative case. No NativeAOT rebuild, full repository test or host numbering/layout acceptance was performed.
