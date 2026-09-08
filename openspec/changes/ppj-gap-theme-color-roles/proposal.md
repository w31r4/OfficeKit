# Explicit authored theme color roles

## Why

F-15 still leaves the six non-accent `a:clrScheme` roles implicit: authored
compilation currently fixes dark/light and hyperlink colors even though the
PPJ theme already has explicit accent and font owners. That leaves no PPJ
field for a presentation author to declare those native roles.

## What Changes

- Add authored `design.theme.colorRoles` with `dark1`, `light1`, `dark2`,
  `light2`, `hyperlink`, and `followedHyperlink` RGB fields.
- Lower the fields to the matching native DrawingML color-scheme roles.
- Keep the existing clean-room defaults for omitted roles and preserve the
  legacy `design.theme.colors` behavior.
- Reject the new field in source-bound PPJ rather than claiming that an
  imported theme part is writable.
- Add a focused native XML, embedded-recovery, and source-bound rejection
  experiment.

## Capabilities

### New Capabilities

- `ppj-authored-theme-color-roles`: Explicit bounded authored ownership of the
  six non-accent theme color roles.

### Modified Capabilities

None.

## Impact

- PPJ v1 schema, generated PPJ reference, the versioned artifact proto,
  authored theme lowering, semantic source-bound validation, coverage/backlog
  wording, and one native test.
- No Office wire-version change.
- Color transforms, alpha, effect schemes, full imported theme XML editing,
  inheritance, and host theme behavior remain outside this slice.
