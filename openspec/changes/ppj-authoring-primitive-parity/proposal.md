## Why

PPJ already exposes more presentation state than its old short authoring notes
suggested, but a direct comparison with a mature finite presentation DSL found
several real usability gaps mixed together with capabilities that already
exist under different names. Adding fields from a checklist without checking
the current compiler would duplicate working state and make the language less
coherent.

## What Changes

- Record a current-schema, compiler-grounded primitive inventory.
- Complete the two missing DrawingML preset geometries.
- Add bounded table style inheritance so Agents do not have to expand every
  repeated cell style by hand.
- Complete the finite Sankey alignment and named-node color controls without
  changing its deterministic vector lowering.
- Document connector state as the existing PPJ line primitive.
- Keep named icons, formulas and native chart-axis/plot options as separately
  scoped follow-ups that require an asset/compiler or wire design.

## Capabilities

### New Capabilities

- `ppj-authoring-primitive-parity`: A first verified authoring-parity batch for
  geometry, compact table styling and deterministic vector Sankey control.

### Modified Capabilities

None. The PPJ schema ID and Office wire version remain unchanged.

## Impact

PPJ schema, preset geometry registry, authored table and Sankey lowering,
generated language reference, visual guidance, coverage and one existing
comprehensive authored-PPJ contract are affected.
