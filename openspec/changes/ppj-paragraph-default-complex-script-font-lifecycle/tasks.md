## 1. Paragraph default complex-script font

- [x] 1.1 Implement exact field authority, presence mutation and targeted a:cs writing; verify assignment/removal/restoration, other-font preservation, name limits and unmodeled-node rejection in shared native fixtures.
- [x] 1.2 Update schema/Help/registry/references and backlog evidence; verify related native tests, generated metadata, preview input/capability, portability/reference sync and strict OpenSpec checks.


Verification (2026-09-10): related native selection passed 177/177, zero skipped,
SDK 8.0.128. Filter covers PpjTextBodyPropertyLifecycleTests, TextParagraph,
DefaultRun and TabStops, excluding the previously clean-baseline-reproduced
ParagraphDefaultRunPropertiesAuthorImportEditAndDeleteWhilePreservingUnknownStyle.
Four lifecycle and two unmodeled complex-script font cases extend the shared
fixtures; assignment/removal/restoration, sole-font wrapper removal, other
script-font preservation, name bounds, authority and native XML/ZIP preservation
pass. Generated reference/matrix, preview input/capability, portability (255
files), reference sync (333 files), strict OpenSpec and whitespace checks pass.
No wire change, AOT rebuild or host font-substitution acceptance.
