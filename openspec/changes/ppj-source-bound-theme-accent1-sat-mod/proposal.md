## Why

The PPJ projection already preserves a narrow set of source-owned accent1 color transforms, but a direct `a:satMod` leaf is still opaque. Exposing this bounded field closes one independent F-15 transform gap while keeping the imported theme graph source-bound and auditable.

## What Changes

- Project a strict imported accent1 `a:satMod` leaf as `design.theme.accentTransforms.accent1.satMod`.
- Issue the field-qualified `setThemeAccent1SatMod` capability for that projection.
- Compile an edited fraction back to only the existing direct `a:satMod/@val` token in the owning ThemePart.
- Reject missing or ambiguous ThemePart ownership, transformed or expanded color topology, unsupported siblings, deletions, combined edits, capability tampering, and out-of-range values.
- Add focused native regression coverage and update the PPJ schema, capability registry, coverage docs, and presentation Skill reference.

## Capabilities

### New Capabilities

- `presentation-theme-accent1-sat-mod`: Source-bound projection and source-preserving compilation of one direct accent1 saturation modulation transform.

### Modified Capabilities

## Impact

The native PPTX projector, compiler, and codec gain one bounded source-bound transform path. The public PPJ schema and capability registry gain one operation and field. Existing protobuf messages and protocol versions are unchanged; authored theme transforms already contain the saturation modulation field.
