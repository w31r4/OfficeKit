## Why

K-05 still describes the host style hierarchy in prose, while the executable
`stylePrecedence` profile only knows a small set of combined sources. That
leaves PPJ authors unable to state, and the compiler unable to prove, the
bounded distinction between layout/master defaults, named styles, element
styles, paragraph defaults, and run overrides.

## What Changes

- Add explicit PPJ text-style sources for `layout`, `master`, `styleRef`,
  `element`, `paragraph`, and `run`, while keeping the legacy `inline` source.
- Add bounded direct text-style fields on text/placeholder elements and on
  layout/master definitions so the hierarchy has concrete PPJ owners.
- Resolve the five scalar text fields `size`, `bold`, `italic`, `font`, and
  `fontFamily` using the declared first-hit order during authored compilation.
- Make the same source vocabulary visible to the read-only grammar review and
  reject unsupported/ambiguous declarations through schema and semantic checks.
- Add one small authored → PPTX → PPJ experiment covering fallback through the
  hierarchy and document the authored-only boundary.

## Capabilities

### New Capabilities

- `ppj-formal-style-precedence`: Bounded, typed host-style precedence for
  authored text-run scalar styles.

### Modified Capabilities

None.

## Impact

- PPJ schema and presentation reference documentation.
- Native PPJ parser/compiler and the JavaScript grammar review evaluator.
- Focused NativeAOT codec regression coverage and the PPJ gap/coverage ledgers.
- No protobuf or wire-version change; source-bound style closure remains
  fail-closed until a native owner is proven.
