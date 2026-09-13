## Why

The PPJ projection now preserves direct accent1 tint, shade, luminance, alpha, saturation, and red modulation transforms, but a direct `a:redOff` leaf is still opaque. Exposing this bounded signed channel offset closes the next independent F-15 transform gap while keeping the imported theme graph source-bound and auditable.

## What Changes

- Project a strict imported accent1 `a:redOff` leaf as `design.theme.accentTransforms.accent1.redOff`.
- Issue the field-qualified `setThemeAccent1RedOff` capability for that projection.
- Compile an edited signed fraction back to only the existing direct `a:redOff/@val` token in the owning ThemePart.
- Reject missing or ambiguous ThemePart ownership, transformed or expanded color topology, unsupported siblings, deletions, combined edits, capability tampering, and out-of-range values.
- Add focused native regression coverage and update the PPJ schema, capability registry, coverage docs, and presentation Skill reference.

## Capabilities

### New Capabilities

- `presentation-theme-accent1-red-off`: Source-bound projection and source-preserving compilation of one direct accent1 red-channel offset transform.

### Modified Capabilities

## Impact

The native PPTX projector, compiler, and codec gain one bounded source-bound transform path. The public PPJ schema and capability registry gain one operation and field. Existing protobuf messages and protocol versions are unchanged; authored theme transforms already contain the red offset field.
