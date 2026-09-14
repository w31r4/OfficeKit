## Why

PPJ already lowers authored `alphaOff` values and has source-bound owners for
accent1 plus accent4 tint, shade, luminance modulation, luminance offset, and
alpha modulation. A direct imported accent4 `a:alphaOff` remains opaque,
leaving one concrete F-15 transform field without a safe edit path.

## What Changes

- Project a canonical shared ThemePart whose accent4 direct RGB leaf contains
  exactly one direct `a:alphaOff/@val` child as
  `design.theme.accentTransforms.accent4.alphaOff`, using a -1..1 fraction.
- Issue `setThemeAccent4AlphaOff` for that field and compile one authorized
  change by rewriting only the existing `a:alphaOff/@val` token.
- Convert the PPJ fraction to DrawingML thousandths with away-from-zero
  rounding and recover it on a fresh projection.
- Preserve no-op bytes, other theme roles, the accent4 RGB value, and
  unsupported topology; reject deletion, siblings, combined edits, bad ranges,
  and capability tampering.
- Add the schema, semantic capability map, registry, generated/reference docs,
  focused native regression, and OpenSpec evidence.

## Capabilities

### New Capabilities

- `presentation-theme-accent4-alphaoff`: Strict source-bound projection and
  source-preserving editing of one direct accent4 `a:alphaOff` leaf.

### Modified Capabilities

None.

## Impact

The native PPTX importer, PPJ projector/compiler, semantic capability map, PPJ
schema, capability registry, generated presentation capability matrix, and
presentation Skill reference gain one additive bounded field. No wire-version
or host-rendering behavior changes.
