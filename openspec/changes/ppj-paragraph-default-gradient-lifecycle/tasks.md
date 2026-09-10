## 1. Paragraph default gradient

- [x] 1.1 Implement exact gradient authority, targeted fill updates, color transitions and gradient opacity/angle normalization; verify lifecycle, wrappers, ordered stops, tokens, neighboring source state, missing authority and invalid/unmodeled paint with focused fixtures and related gradient regressions.
- [x] 1.2 Synchronize schema/Help/registry, focused text reference, generated metadata and backlog evidence; verify explicit preview limits, portability/reference sync, generated checks and strict OpenSpec validation.

## Evidence

- Ten new gradient cases plus four related color cases passed in the first
  focused run; the only failure was the stale percentage-coordinate test.
- The reader already supported centered 50% coordinates in 27e138a4.
  Corrected that test and retained a noncentered 25% negative.
- Related native filter (body lifecycle, TextParagraph, DefaultRun, TabStops,
  Highlight and Gradient) passed 248/248, zero skipped on SDK 8.0.128.
  The documented
  `ParagraphDefaultRunPropertiesAuthorImportEditAndDeleteWhilePreservingUnknownStyle`
  baseline exclusion remains.
- Strengthened the same-name grammar/theme color fixture afterward;
  `FullyQualifiedName~ParagraphDefaultScalarPresencePreservesOtherState&DisplayName~gradient`
  passed 4/4 with the final fixture.
- Preview input/capability and portability/reference-sync checks pass.
  Gradient diagnostics address angle/kind and individual stop fields.
- Text-gradient preview remains partial; no wire change, NativeAOT rebuild
  or host-render acceptance.

- The generated PPJ reference was regenerated and checked against the staged
  schema/registry; the capability matrix and strict OpenSpec checks pass.
