# Change: Complete PPJ preset geometry and adjusted picture masks

## Why

PPJ currently exposes only 49 DrawingML preset geometries even though the
native format and OfficeKit codec recognize a much larger finite vocabulary.
The same incomplete registry also prevents an authored or imported picture
mask from retaining its preset adjustment values. This makes the public
language look and behave materially poorer than the underlying runtime.

## What Changes

- Expand the PPJ shape preset vocabulary to the complete non-connector
  DrawingML preset catalog supported by the pinned Office schema.
- Keep one checked-in profile registry for the PPJ name, native token, ordered
  adjustment guides, defaults, and Agent-facing parameter labels.
- Use the same registry for ordinary shapes and picture masks.
- Compile, project, and capability-edit complete literal picture-mask
  adjustments without exposing raw guide names or formulas.
- Generate the complete shape and mask reference from that registry.

## Impact

- Affected specs: `ppj-preset-geometry-catalog`
- Affected code: PPJ schema/models/validator/compiler/projector, additive wire-v2
  picture state, preset geometry codec, picture codec, generated bindings,
  Presentation Skill reference, capability registry, one integrated test
- Unchanged boundaries: connector topology remains the connector element;
  custom picture-mask paths and noncanonical formula graphs remain source-owned
