## 1. Paragraph default inner shadow

- [x] 1.1 Implement paragraph schema, optional construction/projection, exact authority and direct node patching; verify assignment/removal/restoration, optional geometry/opacity, tokens, mixed effects, original-source preservation and rejected source/input cases with focused and related native tests.
- [x] 1.2 Synchronize Help/registry, focused text reference, generated metadata and backlog evidence; verify explicit preview diagnostics, preserved text-style constraints, portability/reference sync, generated checks and strict OpenSpec validation.

Validation: 307 related native cases passed, zero skipped. Filter:
`(PpjTextBodyPropertyLifecycleTests|TextParagraph|DefaultRun|TabStops|Highlight|Gradient|Glow|InnerShadow|TextReflection|TextSoftEdge|PpjV1ValidatesAndExpandsCanonicalPresentationProgram)`, using FullyQualifiedName substring matches and the existing exclusion `ParagraphDefaultRunPropertiesAuthorImportEditAndDeleteWhilePreservingUnknownStyle`.

Preview input assessment, capability coverage, maintainer check, capability matrix check, Skill portability, reference-sync, Claude plugin and strict OpenSpec validation passed. The full reference-skills test stopped because pdftoppm is unavailable. NativeAOT was not rebuilt; preview remains partial and host display is unverified.
