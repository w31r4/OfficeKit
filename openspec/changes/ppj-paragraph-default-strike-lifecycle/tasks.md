## 1. Paragraph default strike

- [x] 1.1 Implement exact field authority and targeted enum presence writing; verify boolean aliases/canonical strings, cancellation/deletion/restoration, wrappers, invalid/missing authority, direct run/unknown XML preservation and unmodeled-source handling in focused native fixtures.
- [x] 1.2 Synchronize schema/Help/registry/references and backlog evidence; verify related native tests, generated metadata, preview input/capability, portability/reference sync and strict OpenSpec checks.


Completion evidence (2026-09-10): exact defaultText.strike authority is
projected, schema/semantic-validated and independently lowered. Shared
text/shape fixtures verify boolean aliases and canonical strings,
cancellation/removal/restoration, field/wrapper deletion, invalid enum and
missing authority, direct run overrides, original text and non-target XML/ZIP.
A futureStrike native fixture proves no-op identity, replacement rejection
and preservation during unrelated bold assignment/removal. The strengthened
fixture uses valid underline=single for unsupported-field rejection and
adds a single-strike → explicit cancellation → omission chain.

SDK 8.0.128 related filter
`(FullyQualifiedName~PpjTextBodyPropertyLifecycleTests|FullyQualifiedName~TextParagraph|FullyQualifiedName~DefaultRun|FullyQualifiedName~TabStops)&FullyQualifiedName!~ParagraphDefaultRunPropertiesAuthorImportEditAndDeleteWhilePreservingUnknownStyle`
passes 207/207, zero skipped. The documented whole-default-style baseline
exclusion remains unchanged. Maintenance/matrix, preview input/capability,
portability (255 files), reference sync (333 files), strict OpenSpec and
whitespace checks pass. No wire change, NativeAOT rebuild or host acceptance.
