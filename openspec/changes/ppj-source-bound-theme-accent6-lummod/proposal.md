# Source-bound accent6 lumMod

## Why

PPJ already has authored `lumMod` lowering and source-bound owners for
accent1, plus the direct accent6 `tint` and `shade` leaves. A direct imported
`a:lumMod` below accent6 is still opaque. A strict single-leaf accent6 owner
closes one concrete F-15 theme transform gap while keeping the rest of the
imported theme graph source-owned.

## What Changes

- Project a canonical imported presentation whose shared ThemePart contains an
  accent6 direct RGB leaf with exactly one direct `a:lumMod/@val` child as
  `design.theme.accentTransforms.accent6.lumMod`, expressed as a 0..1 fraction.
- Issue `setThemeAccent6LumMod` for that field and compile one authorized
  change by rewriting only the existing `a:lumMod/@val` token.
- Convert the PPJ fraction to DrawingML thousandths with away-from-zero
  rounding and recover the fraction on a fresh projection.
- Keep no-op bytes, other theme roles, the accent6 RGB value, and unsupported
  topology source-owned; reject deletion, sibling or combined edits, bad
  ranges, and capability tampering.
- Add the schema, semantic capability map, registry, generated/reference docs,
  focused native regression, and OpenSpec evidence.

## Capabilities

### New Capabilities

- `presentation-theme-accent6-lumMod`: Strict source-bound projection and
  source-preserving editing of one direct accent6 `a:lumMod` leaf.

### Modified Capabilities

None.

## Impact

The native PPTX importer, PPJ projector/compiler, semantic capability map, PPJ
schema, capability registry, generated presentation capability matrix, and
presentation Skill reference gain one additive bounded field. No wire-version
or host-rendering behavior changes.
