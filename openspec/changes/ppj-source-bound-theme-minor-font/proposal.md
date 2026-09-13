## Why

Source-bound PPJ now exposes and edits a canonical theme's major Latin typeface,
but the already modeled minor Latin field remains read-only even when the same
ThemePart owns it. Closing that adjacent F-15 field provides a small, auditable
font-scheme edit without claiming theme fallback or rewriting the graph.

## What Changes

- Keep projecting `a:minorFont/a:latin/@typeface` as `design.theme.fontScheme.minor` when the canonical major/minor Latin scheme is present.
- Issue a separate hash-bound `setThemeMinorFont` capability for `fontScheme.minor` alongside the existing major field capability.
- Permit one Latin font slot per source-bound compile and patch only the selected minor typeface in the owning ThemePart.
- Preserve major, East Asian, complex-script, descendants, relationships, and other package members; reject deletion, unowned slots, combined major/minor edits, and capability tampering.

## Capabilities

### New Capabilities

- `presentation-theme-minor-font`: Read and source-bound edit of one canonical imported presentation theme minor Latin typeface.

### Modified Capabilities

None.

## Impact

The PPJ theme capability vocabulary, source-bound semantic/compiler path, and
PPTX ThemePart writer gain a minor Latin field owner. The existing
`PresentationThemeArtifact.minor_font_family` wire field is reused; no protobuf
field or protocol version changes. A focused native regression covers exact
no-op bytes, capability authority, one changed ThemePart, non-target package
preservation, Open XML validity, and second projection.
