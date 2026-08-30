## Why

PPJ rich text currently contains only literal text runs. Scientific, technical
and financial presentations therefore have to fake equations with Unicode,
split text boxes or image assets. PowerPoint has a native editable Office Math
inline (`a14:m` with OMML), so a presentation DSL should preserve the semantic
formula instead of rasterizing it.

## What Changes

- Add a typed formula run beside ordinary text runs in PPJ rich text.
- Accept one bounded, documented LaTeX subset and reject unsupported commands
  before PPTX writing.
- Compile the finite formula AST directly in NativeAOT C# to editable native
  Office Math inside a DrawingML paragraph.
- Recover the exact authored LaTeX through the existing embedded PPJ snapshot.
- Recognize OfficeKit's canonical Office Math graph for post-write semantics,
  but never guess LaTeX from arbitrary third-party OMML.
- Teach the Agent when to use a formula and document the supported grammar.

## Capabilities

### New Capabilities

- `ppj-inline-formula-primitive`: bounded LaTeX formula runs compiled to native
  editable PowerPoint Office Math.

### Modified Capabilities

None. The PPJ schema ID and Office wire version remain unchanged.

## Impact

The PPJ schema/model/validator/compiler, additive wire-v2 text inline, native
Presentation text codec, generated language manual, text guidance, coverage
and one existing authored PPJ contract are affected.
