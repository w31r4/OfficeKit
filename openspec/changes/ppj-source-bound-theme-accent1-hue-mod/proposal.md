## Why

The PPJ theme transform surface already supports authored `hueMod`, but a strict imported accent1 hue modulation is still opaque. Exposing this one direct transform closes the next independent F-15 field while keeping the imported theme graph source-bound and auditable.

## What Changes

- Project a strict imported accent1 `a:hueMod` leaf as `design.theme.accentTransforms.accent1.hueMod`.
- Issue the field-qualified `setThemeAccent1HueMod` capability for that projection.
- Compile an edited `0..1` fraction back to only the existing direct `a:hueMod/@val` token in the owning ThemePart.
- Reject missing or ambiguous ThemePart ownership, transformed or expanded color topology, unsupported siblings, deletions, combined edits, capability tampering, and out-of-range values.
- Add focused native regression coverage and update the PPJ schema, capability registry, coverage docs, and presentation Skill reference.

## Capabilities

### New Capabilities

- `presentation-theme-accent1-hue-mod`: Source-bound projection and source-preserving compilation of one direct accent1 hue modulation transform.

### Modified Capabilities

## Impact

The native PPTX projector, compiler, and codec gain one bounded source-bound transform path. The public PPJ schema and capability registry gain one operation and field. Existing protobuf messages and protocol versions are unchanged; authored theme transforms already contain the hue modulation field.
