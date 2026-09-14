## Why

PPJ already lowers authored `lumOff` values and has source-bound owners for
accent1 plus accent3 tint, shade, and luminance modulation. A direct imported
accent3 `a:lumOff` remains opaque, leaving one concrete F-15 transform field
without a safe edit path.

## What Changes

- Project a canonical shared ThemePart whose accent3 direct RGB leaf contains
  exactly one direct `a:lumOff/@val` child as
  `design.theme.accentTransforms.accent3.lumOff`, using a -1..1 fraction.
- Issue `setThemeAccent3LumOff` for that field and compile one authorized
  change by rewriting only the existing `a:lumOff/@val` token.
- Convert the signed PPJ fraction to DrawingML thousandths with away-from-zero
  rounding and recover it on a fresh projection.
- Preserve no-op bytes, other theme roles, the accent3 RGB value, and
  unsupported topology; reject deletion, siblings, combined edits, bad ranges,
  and capability tampering.
- Add the schema, semantic capability map, registry, generated/reference docs,
  focused native regression, and OpenSpec evidence.

## Capabilities

### New Capabilities

- `presentation-theme-accent3-lumoff`: Strict source-bound projection and
  source-preserving editing of one direct accent3 `a:lumOff` leaf.

### Modified Capabilities

None.

## Impact

The native PPTX importer, PPJ projector/compiler, semantic capability map, PPJ
schema, capability registry, generated presentation capability matrix, and
presentation Skill reference gain one additive bounded field. No wire-version
or host-rendering behavior changes.
