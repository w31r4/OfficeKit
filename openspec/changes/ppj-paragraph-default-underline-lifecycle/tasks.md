## 1. Paragraph default underline

- [x] 1.1 Implement exact field authority and targeted native underline writing; verify tokens/aliases, explicit none, deletion/restoration and wrappers, invalid/missing authority, original text/run/XML/ZIP preservation and unmodeled token/effect preservation with focused lifecycle fixtures.
- [x] 1.2 Synchronize schema, Help, registry, references and backlog evidence; verify related native tests, generated metadata, preview input/capability, portability/reference sync and strict OpenSpec checks.


Completion evidence (2026-09-10): independent underline authority is projected,
validated, masked and lowered to the existing optional wire field. Shared
text/shape fixtures cover all 18 native tokens and single/double aliases,
explicit none versus deletion, restoration and wrapper removal, invalid
type/token, missing authority, direct run overrides and non-target XML/ZIP.
Three source fixtures preserve an unknown token and uFillTx with/without u
during no-op and unrelated bold assignment/removal; replacement rejects.
The writer shares the existing uFillTx/uFill/uLnTx/uLn predicate.

Initial run: 210 passed, four failures from assuming fresh PPJ uses native
sng/dbl. Source inspection confirms the established single/double projection.
Specs/docs/tests now retain that contract and assert native and projected
values separately. Final SDK 8.0.128 related filter
`(FullyQualifiedName~PpjTextBodyPropertyLifecycleTests|FullyQualifiedName~TextParagraph|FullyQualifiedName~DefaultRun|FullyQualifiedName~TabStops)&FullyQualifiedName!~ParagraphDefaultRunPropertiesAuthorImportEditAndDeleteWhilePreservingUnknownStyle`
passes 214/214, zero skipped; the known baseline exclusion is unchanged.
Maintenance/matrix, preview input/capability, portability (255 files),
reference sync (333 files), strict OpenSpec and whitespace checks pass.
No wire change, NativeAOT rebuild or host acceptance.
