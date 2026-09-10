## 1. Paragraph default reflection

- [x] 1.1 Connect all optional reflection values, exact field authority and targeted native patching; verify assignment/removal/restoration, empty effects, tokens, positions/transforms, source preservation and rejected input/source cases through focused and related native regressions.
- [x] 1.2 Synchronize Help, registry, focused text guidance, generated metadata and backlog; verify preview diagnostics, maintainer/matrix checks, portability/reference-sync and strict OpenSpec validation.

Native evidence: 323/324 related cases initially passed, zero skipped. The existing full-span source-leaf fixture was updated to explicitly request startPosition 0/endPosition 1; its original assertions remain. That case and the 14 new lifecycle cases then passed 15/15, zero skipped. The maximum-distance fixture also exercises full-span leaf projection, while scale endpoints remain a separate value case.

Related filter uses FullyQualifiedName substring matches:
`(PpjTextBodyPropertyLifecycleTests|TextParagraph|DefaultRun|TabStops|Highlight|Gradient|Glow|InnerShadow|Reflection|TextSoftEdge|PpjV1ValidatesAndExpandsCanonicalPresentationProgram)` with the existing exclusion `ParagraphDefaultRunPropertiesAuthorImportEditAndDeleteWhilePreservingUnknownStyle`.

Preview input assessment, preview capability coverage, generated maintainer/matrix checks, Skill portability, reference-sync and strict OpenSpec validation passed. NativeAOT was not rebuilt; preview remains partial and host appearance is unverified.
