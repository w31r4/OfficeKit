## Why

F-15 names the PowerPoint theme font scheme as a remaining gap, but PPJ has
no explicit owner for it. The authored compiler currently treats the first
two entries of `design.fonts` as major and minor theme fonts, so the native
`a:fontScheme` output depends on an undocumented array convention.

## What Changes

- Add authored `design.theme.fontScheme.major` and `.minor` family fields.
- Lower those fields to the native major/minor font roles in `a:fontScheme`.
- Keep the existing `design.fonts` first/second fallback when the explicit
  scheme is absent.
- Reject the new field in source-bound PPJ rather than inventing a writable
  theme-part capability.
- Add a focused native XML and embedded recovery experiment.

## Capabilities

### New Capabilities

- `ppj-authored-theme-font-scheme`: Explicit bounded major/minor theme font
  ownership for source-free PPJ.

### Modified Capabilities

None.

## Impact

- PPJ v1 schema, generated PPJ reference, authored theme lowering, semantic
  source-bound validation, coverage/backlog wording, and one native test.
- No protobuf or Office wire-version change.
- Full imported `theme1.xml` editing, per-script fallback, font embedding, and
  host font substitution remain outside this slice.
