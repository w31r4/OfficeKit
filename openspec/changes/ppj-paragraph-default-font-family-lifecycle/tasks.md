## 1. Paragraph default font family

- [x] 1.1 Implement field authority, presence changes and targeted Latin font writing; verify lifecycle, bounds, source preservation and unmodeled-font rejection in focused native regressions.
- [x] 1.2 Update schema/Help/registry/references and backlog evidence; verify related native tests, generated metadata, preview input/capability, portability/reference sync and strict OpenSpec checks.


Verification (2026-09-10): related native selection passed 165/165, zero skipped,
SDK 8.0.128; filter covers PpjTextBodyPropertyLifecycleTests, TextParagraph,
DefaultRun and TabStops, with the previously clean-baseline-reproduced
ParagraphDefaultRunPropertiesAuthorImportEditAndDeleteWhilePreservingUnknownStyle
explicitly excluded. Added four font lifecycle cases and two unmodeled-font
cases. First execution caught two wrapper-removal fixtures that also contained
authored a:ea; corrected the source-only fixture to contain only a:latin,
retaining all multi-font preservation assertions. A fixture constructor error
was fixed by assigning parent InnerXml. Generated reference/matrix, preview
input/capability, portability (255 files), reference sync (333 files), strict
OpenSpec and whitespace checks pass. No wire/AOT/host acceptance changes.
