## Why

OfficeKit already authors, imports, validates, and edits all 21 ECMA-376 base
slide-transition effects. PPJ currently exposes only fade, push, wipe, and
Morph, and it omits speed, orientation, effect-specific direction, through-
black behavior, wheel spokes, and click/timed advance. The public language is
therefore materially narrower than the native compiler it is meant to govern.

## What Changes

- Expand PPJ's transition union to the complete bounded base-effect catalog.
- Preserve effect-specific direction, orientation, speed, through-black,
  spokes, duration, and click/timed advance state.
- Lower authored PPJ transitions through the existing canonical Presentation
  transition wire contract.
- Project supported third-party transitions into PPJ and allow source-bound
  local transition edits when the existing codec issues the capability.
- Keep Morph separate, adjacent-page-bound, and incompatible with base-only
  properties.
- Synchronize generated PPJ guidance, motion guidance, coverage, and one
  existing comprehensive PPJ contract.

## Capabilities

### New Capabilities

- `ppj-base-transition-parity`: Complete authored and safely source-bound PPJ
  representation of the existing 21-effect base transition vocabulary.

### Modified Capabilities

None. The PPJ schema ID and Office wire protocol version remain unchanged.

## Impact

- PPJ schema/model validation, authored lowering, PPTX projection,
  source-bound diff lowering, generated documentation, and one existing PPJ
  contract are affected.
- The protobuf already carries the required state, so no wire change is
  required.
- No raw OOXML, arbitrary transition extension, sound, or automatic show
  controller is introduced.
