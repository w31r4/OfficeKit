## Why

Direct rich-text runs already project bounded outer-shadow geometry, but the imported RGB shadow color is still absent from the PPJ native-leaf surface. A narrow color field closes the next documented F-03 sub-gap while preserving the source package through one attribute splice.

## What Changes

- Add `textShadowColorRgb` for `run.style.shadow.color`, using the direct RGB token.
- Project only a strict direct rich-text `a:rPr/a:effectLst/a:outerShdw` owner with one `srgbClr` child and no transforms or sibling effects.
- Replace only `outerShdw/srgbClr/@val` during source-bound edits, preserving alpha, geometry, topology, and unrelated ZIP parts.
- Keep theme-colored, malformed, transformed, and compound effect graphs source-owned or fail closed.
- Update the PPJ schema, capability registry, generated capability matrix, presentation Skill references, coverage/backlog notes, and focused regression evidence.

## Capabilities

### New Capabilities

- `ppj-source-bound-text-shadow-color`: bounded authored and source-bound direct rich-text outer-shadow RGB color with token-preserving edit and reprojection.

### Modified Capabilities

- None.

## Impact

- Native PPJ leaf projection and source-bound edit proof/read/patch dispatch.
- PPJ schema, capability registry, generated capability matrix, presentation Skill, coverage, and backlog documentation.
- A focused codec regression; no protocol version or host-PowerPoint behavior change.
