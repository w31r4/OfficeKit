# Explicit authored theme accent colors

## Why

F-15 names the PowerPoint theme color scheme as a remaining gap, while the
authored compiler currently derives the first six accent colors from the
position of entries in `design.theme.colors`. That convention does not give a
PPJ author an explicit owner for the native `accent1` through `accent6` roles.

## What Changes

- Add authored `design.theme.accentColors.accent1` through `.accent6` RGB
  fields.
- Lower those fields to the native `a:accent1Color` through `a:accent6Color`
  roles.
- Keep the existing `design.theme.colors` first-six fallback when the explicit
  role map is absent.
- Reject the new field in source-bound PPJ rather than claiming that an
  imported theme part is writable.
- Add a focused native XML and embedded recovery experiment.

## Capabilities

### New Capabilities

- `ppj-authored-theme-accent-colors`: Explicit bounded authored ownership of
  the six native theme accent colors.

### Modified Capabilities

None.

## Impact

- PPJ v1 schema, generated PPJ reference, authored theme lowering, semantic
  source-bound validation, coverage/backlog wording, and one native test.
- No protobuf or Office wire-version change.
- Non-accent theme roles, color transforms, effect schemes, full imported
  `theme1.xml` editing, and host theme behavior remain outside this slice.
