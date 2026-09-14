## Why

PPJ already lowers authored `redOff` values and has source-bound owners for
accent1 plus accent5 tint, shade, luminance, alpha, and saturation transforms.
A direct imported accent5 `a:redOff` remains opaque, leaving one concrete F-15
transform field without a safe edit path.

## What Changes

- Project a canonical shared ThemePart whose accent5 direct RGB leaf contains
  exactly one direct `a:redOff/@val` child as
  `design.theme.accentTransforms.accent5.redOff`, using a -1..1 fraction.
- Issue `setThemeAccent5RedOff` for that field and compile one authorized
  change by rewriting only the existing `a:redOff/@val` token.
- Convert the PPJ fraction to DrawingML thousandths with away-from-zero
  rounding and recover it on a fresh projection.
- Preserve no-op bytes, other theme roles, the accent5 RGB value, and
  unsupported topology; reject deletion, siblings, combined edits, bad ranges,
  and capability tampering.
- Add the schema, semantic capability map, registry, generated/reference docs,
  focused native regression, and OpenSpec evidence.

## Capabilities

### New Capabilities

- `presentation-theme-accent5-redoff`: Strict source-bound projection and
  source-preserving editing of one direct accent5 `a:redOff` leaf.

### Modified Capabilities

None.

## Impact

The native PPTX importer, PPJ projector/compiler, semantic capability map, PPJ
schema, capability registry, generated presentation capability matrix, and
presentation Skill reference gain one additive bounded field. No wire-version
or host-rendering behavior changes.
