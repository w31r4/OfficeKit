## Why

PPJ already lowers authored `satMod` values and has source-bound owners for
accent1 plus accent3 tint, shade, luminance, and alpha transforms. A direct
imported accent3 `a:satMod` remains opaque, leaving one concrete F-15
transform field without a safe edit path.

## What Changes

- Project a canonical shared ThemePart whose accent3 direct RGB leaf contains
  exactly one direct `a:satMod/@val` child as
  `design.theme.accentTransforms.accent3.satMod`, using a 0..1 fraction.
- Issue `setThemeAccent3SatMod` for that field and compile one authorized
  change by rewriting only the existing `a:satMod/@val` token.
- Convert the PPJ fraction to DrawingML thousandths with away-from-zero
  rounding and recover it on a fresh projection.
- Preserve no-op bytes, other theme roles, the accent3 RGB value, and
  unsupported topology; reject deletion, siblings, combined edits, bad ranges,
  and capability tampering.
- Add the schema, semantic capability map, registry, generated/reference docs,
  focused native regression, and OpenSpec evidence.

## Capabilities

### New Capabilities

- `presentation-theme-accent3-satmod`: Strict source-bound projection and
  source-preserving editing of one direct accent3 `a:satMod` leaf.

### Modified Capabilities

None.

## Impact

The native PPTX importer, PPJ projector/compiler, semantic capability map, PPJ
schema, capability registry, generated presentation capability matrix, and
presentation Skill reference gain one additive bounded field. No wire-version
or host-rendering behavior changes.
