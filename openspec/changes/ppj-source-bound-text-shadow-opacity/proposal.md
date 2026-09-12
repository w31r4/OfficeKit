## Why

Direct rich-text runs already project bounded outer-shadow geometry and can
carry an authored opacity in the semantic shadow object, but an imported
direct alpha token is still missing from the PPJ native-leaf surface. A narrow
opacity field closes the next F-03 sub-gap while preserving the source package
through one attribute splice.

## What Changes

- Add `textShadowOpacityThousandthPercent` for `run.style.shadow.opacity`.
- Project one strict direct rich-text `a:rPr/a:effectLst/a:outerShdw` whose
  single RGB or theme color child already contains one bounded `a:alpha/@val`.
- Replace only that existing alpha token during source-bound edits, preserving
  color, geometry, topology, and unrelated ZIP parts.
- Keep missing, malformed, transformed, compound, and out-of-range effect
  graphs source-owned or fail closed.
- Update the PPJ schema, capability registry, generated capability matrix,
  presentation Skill references, coverage/backlog notes, and focused
  regression evidence.

## Capabilities

### New Capabilities

- `ppj-source-bound-text-shadow-opacity`: bounded authored and source-bound
  direct rich-text outer-shadow alpha with token-preserving edit and
  reprojection.

### Modified Capabilities

- None.

## Impact

- Native PPJ leaf projection and source-bound edit proof/read/patch dispatch.
- PPJ schema, capability registry, generated capability matrix, presentation
  Skill, coverage, and backlog documentation.
- A focused codec regression; no protocol version or host-PowerPoint behavior
  change.
