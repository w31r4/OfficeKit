## Why

F-03 list color still cannot be edited through ordinary source-bound PPJ. Native follow-text color is omitted from PPJ, authored theme colors become RGB, and RGB alpha projection rounds native precision to eight bits.

## What Changes

- Complete the mutually exclusive `bullet.color` / `colorFollowText: true` choice for character, numbered and picture markers, including absence, switching, removal and restoration.
- Retain direct standard theme-token identity and explicit alpha; resolve declared grammar colors and tint/shade through the existing color rules.
- Add `{ rgb: "#RRGGBB", alpha: number }` to bullet color only, preserving native alpha precision where hex RGBA cannot. Existing hex and token forms remain valid.
- Edit only the target color declaration under exact changed-field authority; preserve marker identity, font/size, surrounding XML and ZIP parts. Preserve unknown/malformed source color choices and reject replacement.
- Keep direct RGB preview and explicit theme/follow-text layout limitations discoverable.

## Capabilities

### New Capabilities

- `ppj-bullet-color-lifecycle`: direct marker color choice, precise alpha, native identity, source lifecycle and preservation.

### Modified Capabilities

None.

## Impact

PPJ schema, authored/source compilers, projector/authority, shared bullet style codec, Help, registry/manual/matrix, text guidance and focused tests. The native wire already represents these choices; no protocol or dependency change. Precise RGB alpha may project as an object where older output used a rounded hex string.
