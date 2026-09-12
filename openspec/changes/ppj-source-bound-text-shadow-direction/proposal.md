## Why

Direct rich-text runs already carry authored outer-shadow direction, but an imported run currently exposes only blur and distance as native PPJ leaves. A bounded direction leaf closes the next documented F-03 text-effect gap while preserving the source package through a single attribute splice.

## What Changes

- Add `textShadowDirectionDegrees` for `run.style.shadow.angle`, represented by the native 1/60000-degree token.
- Project only a strict direct rich-text `a:outerShdw/@dir` owner with one RGB/theme color child and no shadow transforms or sibling effects.
- Replace only the existing `dir` token during source-bound edits, preserving the owning SlidePart, effect topology, text runs, and unrelated package parts.
- Keep malformed, out-of-range, transformed, or complex effect lists source-owned/fail-closed.
- Update the PPJ schema, capability registry, generated capability matrix, presentation Skill references, coverage/backlog notes, and focused regression evidence.

## Capabilities

### New Capabilities

- `ppj-source-bound-text-shadow-direction`: bounded authored and source-bound direct rich-text outer-shadow direction with token-preserving edit and reprojection.

### Modified Capabilities

- None.

## Impact

- Native PPJ leaf projection and source-bound edit proof/read/patch dispatch.
- PPJ schema, capability registry, generated capability matrix, presentation Skill, coverage, and backlog documentation.
- A focused codec regression; no protocol version or host-PowerPoint behavior change.
