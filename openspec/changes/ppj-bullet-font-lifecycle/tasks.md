## 1. Direct bullet font choice

- [x] 1.1 Implement schema/native validation, follow-text lowering/projection and exact source mutation with choice-local native apply; verify authored/source add/set/switch/remove/restore, authority, invalid input and unmodeled preservation using focused text/shape marker tests.

## 2. Shared behavior and discoverability

- [x] 2.1 Update Help, schema descriptions, registry, generated manual/matrix, text guidance and F-03 backlog; verify preview diagnostics, related paragraph/list/table/master tests, generated checks, Skill portability/reference sync and strict OpenSpec validation.

Validation (2026-09-10): the initial cases reproduced rejected follow-text input and inconsistent supplementary-character length validation. The completed focused tests pass 4/4, exercising ordinary text/shape lifecycle on character, number and picture markers, authored external URI and imported embedded media, native follow-text import, exact field authority, invalid Unicode/XML and unknown font preservation. Additional original-spelling assertions first failed for both owners, then passed after retaining unchanged color/size nodes. Related paragraph/list/table/master regressions pass 30/30 including these cases, zero skips. Preview retains direct character-font painting and diagnoses follow-text layout unavailable. Help/generated matrix/manual, portability/reference sync and strict OpenSpec checks pass. No full repository test, NativeAOT rebuild or PowerPoint host font/layout acceptance was performed.
