## 1. Paragraph default East Asian font

- [x] 1.1 Add independent field authority, explicit-presence mutation and targeted font writing; verify assignment/removal/restoration, Latin-retained deletion, bounds and unmodeled-node rejection in the shared native fixtures.
- [x] 1.2 Update schema/Help/registry/references and backlog evidence; verify related native tests, generated metadata, preview input/capability, portability/reference sync and strict OpenSpec checks.


Verification (2026-09-10): related native selection passed 171/171, zero skipped,
SDK 8.0.128. Filter covers PpjTextBodyPropertyLifecycleTests, TextParagraph,
DefaultRun and TabStops, excluding the previously clean-baseline-reproduced
ParagraphDefaultRunPropertiesAuthorImportEditAndDeleteWhilePreservingUnknownStyle.
Four lifecycle and two unmodeled East Asian font cases were added to shared
fixtures. Latin-retained deletion proves raw presence bypasses authored fallback;
font-only removal/restoration, name boundaries, authority, native XML and ZIP
preservation pass. Generated reference/matrix, preview input/capability,
portability (255 files), reference sync (333 files), strict OpenSpec and whitespace
checks pass. No wire change, AOT rebuild or host acceptance.
