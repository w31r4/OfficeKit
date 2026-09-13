## Why

Source-bound PPJ now owns the canonical theme's Latin and East Asian faces,
but an existing major complex-script face remains read-only when its ThemePart
slot is known. Closing this adjacent F-15 field adds a truthful script-specific
edit without claiming full font fallback.

## What Changes

- Keep projecting an existing `a:majorFont/a:cs/@typeface` as `design.theme.fontScheme.majorComplexScript` when the canonical Latin scheme and this slot are present.
- Issue a hash-bound `setThemeMajorFontComplexScript` capability for `fontScheme.majorComplexScript`.
- Permit one font-scheme slot per source-bound compile and patch only the selected major complex-script typeface in the owning ThemePart.
- Preserve Latin, East Asian, minor complex-script, descendants, relationships, and other package members; reject deletion, missing ownership, combined edits, and capability tampering.

## Capabilities

### New Capabilities

- `presentation-theme-major-complex-script-font`: Read and source-bound edit of one canonical imported presentation theme major complex-script typeface.

### Modified Capabilities

None.

## Impact

The PPJ theme capability vocabulary, source-bound compiler, projection, and
PPTX ThemePart writer gain a major complex-script field owner. The existing
`PresentationThemeArtifact.major_font_family_complex_script` wire field is
reused; no protobuf field or protocol version changes. A focused native
regression covers exact no-op bytes, capability authority, one changed
ThemePart, non-target package preservation, Open XML validity, and second
projection.
