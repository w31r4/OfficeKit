## Why

The bounded formal text precedence profile covers Latin font ownership but
leaves the existing `fontFamilyEastAsia` field on the old inline/default path.
That makes an explicitly declared precedence rule unable to control the
East Asian typeface and leaves a real rich-text field outside the same owner
chain.

## What Changes

- Add `text.fontFamilyEastAsia` to the bounded formal text precedence targets.
- Resolve run, paragraph, element, named-style, layout, master, theme, and
  default East Asian font values in declared order.
- Preserve the existing Latin-font fallback when no East Asian value resolves.
- Make authored compilation and read-only grammar review report the same
  winning source and value.
- Keep source-bound formal owner declarations fail closed and leave broader
  language/font fallback behavior unchanged.

## Capabilities

### New Capabilities

- `ppj-formal-eastasia-font-owner`: Formal precedence ownership for the
  existing East Asian run typeface field.

### Modified Capabilities

None.

## Impact

- PPJ reference/backlog wording, formal grammar validation, authored text
  lowering, and read-only review evidence.
- One focused round-trip extends the existing formal text test with a theme
  East Asian font winner and default fallthrough.
- No protobuf or Office wire-version change; no native theme scheme or host
  font fallback graph is added.
