## Why

Imported presentations currently expose a placeholder theme name and reject every source-bound theme edit. A bounded theme-name owner gives PPJ one complete, auditable field from the existing F-15 theme gap without pretending to edit the rest of `theme1.xml`.

## What Changes

- Project the existing canonical presentation theme name into `design.theme.name` with a hash-bound native reference.
- Issue a `setThemeName` capability only when all masters share one existing ThemePart with an existing theme name.
- Preserve every other theme child and reject missing, ambiguous, or unsupported theme owners.
- Patch only the ThemePart on a source-bound name edit and verify the name after re-import.

## Capabilities

### New Capabilities

- `presentation-theme-name`: Read and source-bound edit of one canonical imported presentation theme name.

### Modified Capabilities

None.

## Impact

The PPJ schema and presentation projector gain a theme-level native reference. The PPTX importer and source-preserving writer read and patch the canonical `a:theme/@name` attribute. A focused codec regression covers authoring, source projection, capability authority, changed-part scope, no-op bytes, Open XML validity, and second projection. The complete theme color/font/effect graph remains source-owned.
