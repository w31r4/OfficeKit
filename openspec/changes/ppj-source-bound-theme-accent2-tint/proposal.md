# Source-bound accent2 tint

## Why

PPJ already has an authored `tint` transform and a source-bound owner for
accent1, but a direct imported `a:tint` below accent2 is still opaque. A strict
single-leaf accent2 owner closes one concrete F-15 theme transform gap while
keeping the rest of the imported theme graph source-owned.

## What Changes

- Project a canonical imported presentation whose shared ThemePart contains an
  accent2 direct RGB leaf with exactly one direct `a:tint/@val` child as
  `design.theme.accentTransforms.accent2.tint`, expressed as a 0..1 fraction.
- Issue `setThemeAccent2Tint` for that field and compile one authorized change
  by rewriting only the existing `a:tint/@val` token.
- Convert the PPJ fraction to DrawingML thousandths with away-from-zero
  rounding and recover the fraction on a fresh projection.
- Keep no-op bytes, other theme roles, the accent2 RGB value, and unsupported
  topology source-owned; reject deletion, sibling or combined edits, bad
  ranges, and capability tampering.
- Add the schema, semantic capability map, registry, generated/reference docs,
  focused native regression, and OpenSpec evidence.

## Capabilities

### New Capabilities

- `presentation-theme-accent2-tint`: Strict source-bound projection and
  source-preserving editing of one direct accent2 `a:tint` leaf.

### Modified Capabilities

None.

## Impact

The native PPTX importer, PPJ projector/compiler, semantic capability map, PPJ
schema, capability registry, generated presentation capability matrix, and
presentation Skill reference gain one additive bounded field. No wire-version
or host-rendering behavior changes.
