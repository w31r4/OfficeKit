# Source-bound accent2 shade

## Why

PPJ already has authored `shade` lowering and source-bound owners for
accent1 plus the direct accent2 `tint` leaf. A direct imported `a:shade`
below accent2 is still opaque. A strict single-leaf accent2 owner closes one
concrete F-15 theme transform gap while keeping the rest of the imported theme
graph source-owned.

## What Changes

- Project a canonical imported presentation whose shared ThemePart contains an
  accent2 direct RGB leaf with exactly one direct `a:shade/@val` child as
  `design.theme.accentTransforms.accent2.shade`, expressed as a 0..1 fraction.
- Issue `setThemeAccent2Shade` for that field and compile one authorized change
  by rewriting only the existing `a:shade/@val` token.
- Convert the PPJ fraction to DrawingML thousandths with away-from-zero
  rounding and recover the fraction on a fresh projection.
- Keep no-op bytes, other theme roles, the accent2 RGB value, and unsupported
  topology source-owned; reject deletion, sibling or combined edits, bad
  ranges, and capability tampering.
- Add the schema, semantic capability map, registry, generated/reference docs,
  focused native regression, and OpenSpec evidence.

## Capabilities

### New Capabilities

- `presentation-theme-accent2-shade`: Strict source-bound projection and
  source-preserving editing of one direct accent2 `a:shade` leaf.

### Modified Capabilities

None.

## Impact

The native PPTX importer, PPJ projector/compiler, semantic capability map, PPJ
schema, capability registry, generated presentation capability matrix, and
presentation Skill reference gain one additive bounded field. No wire-version
or host-rendering behavior changes.
