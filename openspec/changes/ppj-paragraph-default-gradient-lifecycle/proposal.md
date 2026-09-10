## Why

F-03 paragraph default gradients can be authored and projected, but source-bound structured edits lack independent gradient authority. Complete the field lifecycle beside the existing color lifecycle.

## What Changes

- Allow defaultText.gradient assignment, deletion, gradient-only wrapper removal and restoration on ordinary text/shape paragraphs.
- Preserve existing linear and centered radial profiles, ordered RGB/token stops, alpha and native precision.
- Allow solid/gradient transitions only when both changed fields have exact authority.
- Update only changed paragraphs; retain other paragraphs, direct run paint and unrelated native state.
- Reject unsupported source paint/transforms and invalid gradients without output; synchronize guidance and focused regression evidence.

## Capabilities

### New Capabilities
- `ppj-paragraph-default-gradient-lifecycle`: Independent paragraph default gradient editing and color/gradient transitions.

### Modified Capabilities

None.

## Impact

PPJ field authority, paragraph diff/lowering, native default-run fill writer, schema/Help/registry/references and focused fixtures. Reuse existing gradient wire messages. Other fill graphs and host rendering remain separate capabilities.
