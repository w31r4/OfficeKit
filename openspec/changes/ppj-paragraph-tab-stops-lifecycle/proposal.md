## Why

F-03 tab stops can be assigned, but ordinary PPJ deletion is rejected and a change to one paragraph rewrites every paragraph's tab list. The current reader also accepts partial lists and reads malformed SDK numbers/enums before classifying source content.

## What Changes

- Complete `text.paragraphs[].style.tabStops` add/set/clear/remove/restore for ordinary source-bound text/shape owners. Empty arrays and removal clear the modeled direct list; retain `noTabStops: true` as an explicit removal alias that projects back to absence.
- Keep at most 32 strictly increasing positions after nearest-EMU rounding, with left/center/right/decimal alignment. Align PPJ's nonnegative point range with the native signed-32-bit EMU range, including imported maximum values.
- Treat tabStops/noTabStops as one mutually exclusive choice at each style precedence layer. Require exact authority for every changed property, and only mutate changed paragraphs.
- Preserve semantically unchanged native numeric spelling and alignment omission. Retain malformed, duplicate and unknown tab lists as source content; reject replacements without blocking independent modeled text/style edits.
- Update Help, schema, capability/Agent documentation and focused lifecycle/preview regressions. Preview keeps its explicit tab-layout limitations.

## Capabilities

### New Capabilities

- `ppj-paragraph-tab-stops-lifecycle`: bounded tab-list semantics, native preservation and complete source editing lifecycle.

### Modified Capabilities

None. Main specs are empty; existing tab-stop behavior is documented in code and prior change artifacts.

## Impact

PPJ schema, authored/source compilers, projected field authority, shared text/paragraph codec, focused native/preview tests, Help, registry/manual/matrix, text guidance and F-03 backlog. Existing repeated tab-stop wire and explicit removal flag are reused; no dependency or wire version change. Inherited tab resolution and host text measurement remain separate.
