## Why

Source-bound PPJ now owns the canonical theme's Latin faces and major East
Asian face, but an existing minor East Asian face remains read-only when its
ThemePart slot is known. Closing this adjacent F-15 field adds a truthful
script-specific edit without claiming full font fallback.

## What Changes

- Keep projecting an existing `a:minorFont/a:ea/@typeface` as `design.theme.fontScheme.minorEastAsia` when the canonical Latin scheme and this slot are present.
- Issue a hash-bound `setThemeMinorFontEastAsia` capability for `fontScheme.minorEastAsia`.
- Permit one font-scheme slot per source-bound compile and patch only the selected minor East Asian typeface in the owning ThemePart.
- Preserve Latin, major East Asian, complex-script, descendants, relationships, and other package members; reject deletion, missing ownership, combined edits, and capability tampering.

## Capabilities

### New Capabilities

- `presentation-theme-minor-eastasia-font`: Read and source-bound edit of one canonical imported presentation theme minor East Asian typeface.

### Modified Capabilities

None.

## Impact

The PPJ theme capability vocabulary, source-bound compiler, projection, and
PPTX ThemePart writer gain a minor East Asian field owner. The existing
`PresentationThemeArtifact.minor_font_family_east_asia` wire field is reused;
no protobuf field or protocol version changes. A focused native regression
covers exact no-op bytes, capability authority, one changed ThemePart,
non-target package preservation, Open XML validity, and second projection.
