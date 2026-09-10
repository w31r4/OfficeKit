## 1. Paragraph default highlight

- [x] 1.1 Implement exact field authority, RGB/theme projection and per-paragraph targeted writing; verify assignment/deletion/restoration, wrapper removal, tokens, untouched theme bindings, invalid/missing authority and unmodeled source graph preservation and malformed-source fail-closed behavior in focused fixtures.
- [x] 1.2 Synchronize schema/Help/registry/references and backlog evidence; verify related native tests, generated metadata, preview input/capability, portability/reference sync and strict OpenSpec checks.


Completion evidence (2026-09-10): exact highlight authority, native RGB/scheme
projection, per-paragraph mutation and targeted child writing are implemented.
Shared lifecycle fixtures cover opaque RGB/case/alpha=1, grammar tint/shade,
deletion/restoration/wrappers, invalid colors/alpha and missing authority.
Two text/shape source fixtures verify theme replacement/deletion/restoration,
neighboring theme preservation and grammar precedence (including alpha rejection).
Source fixtures preserve transformed colors, duplicates and unknown elements;
malformed bare character data keeps byte-identical no-op and edit rejection.

First run: 224/227, with two theme/grammar precedence failures and one malformed
text binding failure. Explicit grammar resolution fixed precedence. Raw XML
checks run before typed children; malformed text still cannot be retained by
the SDK, so specs/tests follow fail-closed source binding rather than weaken
the hash check. A separate unknown-element fixture proves preservation.
Six focused highlight cases pass. Final SDK 8.0.128 related filter
`(FullyQualifiedName~PpjTextBodyPropertyLifecycleTests|FullyQualifiedName~TextParagraph|FullyQualifiedName~DefaultRun|FullyQualifiedName~TabStops|FullyQualifiedName~Highlight)&FullyQualifiedName!~ParagraphDefaultRunPropertiesAuthorImportEditAndDeleteWhilePreservingUnknownStyle`
passes 228/228, zero skipped. Existing baseline exclusion remains.
Maintenance/matrix, preview input/capability, portability (255 files),
reference sync (333 files), strict OpenSpec and whitespace checks pass.
Temporary hash diagnostics are removed. No wire change, NativeAOT rebuild
or host acceptance.
