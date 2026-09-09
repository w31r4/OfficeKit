## Why

F-03 paragraph default baseline can be authored and projected but lacks independent structured source-bound editing. Complete its signed percentage and presence lifecycle alongside the existing default scalar fields.

## What Changes

- Enable defaultText.baseline assignment, explicit zero, deletion, baseline-only wrapper removal and restoration for ordinary text/shape paragraphs.
- Retain finite -400..400 percent values with native thousandths of a percent and ties-to-even rounding.
- Preserve other defaults, direct run baseline and unknown source content; reject replacement of unmodeled native baseline.
- Add focused regressions and synchronize field guidance.

## Capabilities

### New Capabilities
- `ppj-paragraph-default-baseline-lifecycle`: Independent direct paragraph default baseline lifecycle.

### Modified Capabilities

None.

## Impact

PPJ schema/authority, paragraph diff/mutation, default-run writer, tests and references. Existing optional FontBaselinePercent suffices; no wire change. Host layout and inherited placeholder defaults remain separate.
