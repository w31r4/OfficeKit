## Why

The formal text precedence profile now names `theme`, but the canonical PPJ
theme has no text-style owner that can provide a value. That leaves the lowest
formal layer unexecutable and makes a theme hit indistinguishable from a
missing value. Add one explicit authored fallback owner so the precedence
chain can be exercised without pretending to edit arbitrary PowerPoint theme
XML.

## What Changes

- Add optional `design.theme.textStyle` as a direct authored text-style owner.
- Resolve its `size`, `bold`, `italic`, `font`, and `fontFamily` values when a
  formal `stylePrecedence` rule declares `theme`.
- Make the read-only reviewer report `theme` as the winning source when it
  supplies the first value, with normal fallthrough to `default`.
- Keep the existing native theme artifact and text-container style semantics;
  authored compilation lowers only the effective run scalar values.
- Reject source-bound programs that carry this new declaration so they do not
  claim source-owned theme style editing.

## Capabilities

### New Capabilities

- `ppj-theme-text-style-owner`: A bounded authored theme text-style fallback
  for the formal PPJ text precedence profile.

### Modified Capabilities

None.

## Impact

- PPJ schema and presentation reference gain one optional theme owner.
- The authored NativeAOT compiler and JavaScript reviewer share the same
  five-field theme lookup.
- A focused C# round-trip and review smoke cover theme hit and missing-theme
  fallback cases.
- No protobuf or Office wire-version change; full `theme1.xml` font/effect
  scheme inheritance remains outside this change.
