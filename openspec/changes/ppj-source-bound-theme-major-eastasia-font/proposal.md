## Why

Source-bound PPJ now owns the canonical theme's Latin major and minor faces,
but an existing major East Asian face remains read-only even when its exact
ThemePart slot is known. Closing this adjacent F-15 field gives PPJ a truthful
script-specific edit without pretending to implement theme fallback.

## What Changes

- Keep projecting an existing `a:majorFont/a:ea/@typeface` as `design.theme.fontScheme.majorEastAsia` when the canonical Latin scheme and this slot are present.
- Issue a hash-bound `setThemeMajorFontEastAsia` capability for `fontScheme.majorEastAsia`.
- Permit one font-scheme slot per source-bound compile and patch only the selected major East Asian typeface in the owning ThemePart.
- Preserve Latin, minor East Asian, complex-script, descendants, relationships, and other package members; reject deletion, missing ownership, combined edits, and capability tampering.

## Capabilities

### New Capabilities

- `presentation-theme-major-eastasia-font`: Read and source-bound edit of one canonical imported presentation theme major East Asian typeface.

### Modified Capabilities

None.

## Impact

The PPJ theme capability vocabulary, source-bound compiler, projection, and
PPTX ThemePart writer gain a major East Asian field owner. The existing
`PresentationThemeArtifact.major_font_family_east_asia` wire field is reused;
no protobuf field or protocol version changes. A focused native regression
covers exact no-op bytes, capability authority, one changed ThemePart,
non-target package preservation, Open XML validity, and second projection.
