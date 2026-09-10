## 1. Paragraph default color

- [x] 1.1 Implement exact color authority and targeted per-paragraph RGB/theme/alpha writing; verify assignment/deletion/restoration, wrappers, tokens/alpha presence, neighboring precision, invalid/missing authority and source paint preservation/rejection in focused fixtures.
- [x] 1.2 Synchronize schema/Help/registry/references and backlog evidence; verify related native tests, generated metadata, preview input/capability, portability/reference sync and strict OpenSpec checks.

## Evidence

- SDK 8.0.128 related filter (body lifecycle, TextParagraph, DefaultRun,
  TabStops and Highlight): 235/237 initially passed; both failures were
  JSON property-order comparisons in the two new theme-color cases.
- After semantic JSON comparison and color/gradient negative fixtures,
  the focused color filter passed 9/9, zero skipped. The other 228 related
  cases passed in the initial run; production code did not change afterward.
- Focused filter:
  `FullyQualifiedName~ParagraphDefaultColor|(FullyQualifiedName~ParagraphDefaultScalarPresencePreservesOtherState&DisplayName~color)`.
- Related filter retains the documented
  `ParagraphDefaultRunPropertiesAuthorImportEditAndDeleteWhilePreservingUnknownStyle`
  baseline exclusion.
- Presentation Skill maintenance, generated capability matrix, preview
  input/capability, portability/reference sync and strict OpenSpec pass.
- No wire change, NativeAOT rebuild or PowerPoint host acceptance for this field.
