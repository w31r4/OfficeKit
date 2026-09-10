## Why

F-03 still lacks independent source-bound editing for paragraph default soft edges. The existing radius field already authors and projects, but changing or deleting it has no field authority and can enter a whole-style rewrite.

## What Changes

- Complete `text.paragraphs[].style.defaultText.softEdge` assignment, deletion and restoration on ordinary text/shape owners, including removal of a wrapper containing only that effect.
- Retain the existing required radius in points, finite 0..1000 with native EMU rounding; explicit zero retains the effect.
- Patch only the direct soft-edge node, preserving other effects, list attributes, runs, neighboring paragraphs and non-target ZIP parts. Unsupported source content remains preserved and rejects replacement.
- Synchronize schema authority, Help, registry, text guidance, generated references and backlog; retain partial preview diagnostics.

## Capabilities

### New Capabilities

- `ppj-paragraph-default-soft-edge-lifecycle`: independent direct paragraph default soft-edge editing and source preservation.

### Modified Capabilities

None.

## Impact

PPJ source compiler/projector/validation, default-run effect patching, focused native lifecycle tests and presentation metadata. No wire change, new dependency or NativeAOT/host acceptance claim.
