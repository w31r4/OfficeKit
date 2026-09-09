## 1. Paragraph default size

- [x] 1.1 Extend size authority, presence mutation and native scalar writing; verify assignment/removal/restoration, precision, denied edits and non-target preservation in the focused lifecycle regression.
- [x] 1.2 Update schema/Help/registry/presentation guidance and backlog evidence; verify related native tests, generated metadata, preview assessment, portability/reference checks and strict OpenSpec validation.


Verification (2026-09-10): related native selection passed 159/159, zero skipped,
SDK 8.0.128. Filter covers PpjTextBodyPropertyLifecycleTests, TextParagraph,
DefaultRun and TabStops, excluding the previously clean-baseline-reproduced
ParagraphDefaultRunPropertiesAuthorImportEditAndDeleteWhilePreservingUnknownStyle.
Initial focused run passed eight booleans but rejected four 0.01pt restorations:
native sz minimum is 100. Corrected the declared range and default-run validator
to 1..768pt; retained 0.01pt among invalid cases and verified both numeric bounds
and nearest-even rounding. Generated reference/matrix, preview input/capability,
portability (255 files), reference sync (333 files), strict OpenSpec and whitespace
checks pass. No protocol change, AOT rebuild or host acceptance.
