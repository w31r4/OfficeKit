## Why

F-03 paragraph default fontFamily is authored and projected but lacks structured source-bound assignment/removal. The existing Latin font writer already defines a simple-node preservation boundary.

## What Changes

- Add independently authorized defaultText.fontFamily assignment, removal and restoration for ordinary text/shape paragraphs.
- Support font-only wrapper deletion while retaining other defaults, script fonts, effects and direct runs.
- Reject replacement of unmodeled Latin font nodes; cover additional attributes and hidden leaf children.
- Extend the shared lifecycle regression and public field guidance.

## Capabilities

### New Capabilities
- `ppj-paragraph-default-font-family-lifecycle`: Direct paragraph default Latin font presence.

### Modified Capabilities

None.

## Impact

Capability schema/projection/validation, source compiler, default-run font writer and presentation references. Existing FontFamily wire presence suffices. Font discovery and host substitution remain outside this field lifecycle.
