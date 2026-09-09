## 1. Paragraph default bold

- [x] 1.1 Implement exact capability, diff and presence lowering; verify text/shape lifecycle, sibling XML/ZIP preservation and authority rejection with focused native tests.
- [x] 1.2 Update schema/Help/registry/reference and backlog evidence; verify generated references/matrix, preview diagnostics, portability/reference sync and strict OpenSpec.

Evidence: direct lifecycle/tab-stop selection 5/5 passed; related native selection 151/151 passed, zero skipped (SDK 8.0.128). The known ParagraphDefaultRunPropertiesAuthorImportEditAndDeleteWhilePreservingUnknownStyle failure was reproduced on clean bbc1065b baseline and excluded from the 151-test selection. Exact-source tab-stop fixture stabilized. Generated reference/matrix, preview input/capability, portability/reference sync and strict OpenSpec passed. Existing wire unchanged; no NativeAOT rebuild or host acceptance.
