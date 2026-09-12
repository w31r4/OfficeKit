## Why

Direct rich-text runs already carry authored outer-shadow values through PPJ, but imported runs do not expose a native leaf for even the bounded blur radius. That leaves a common text effect opaque during source-preserving edits and makes the documented F-03 surface uneven; a single token-bound blur field is the smallest useful closure.

## What Changes

- Add the `textShadowBlurRadiusEmu` native leaf for `run.style.shadow.blur`.
- Project only a strict direct rich-text `a:outerShdw/@blurRad` owner with one RGB/theme color child and no shadow transforms or sibling effect nodes.
- Source-bound edits replace only the existing `blurRad` token and preserve the owning SlidePart's other XML, relationships, media, text runs, and neighboring effects.
- Reproject authored and edited packages with the typed shadow blur value; keep malformed, transformed, or complex effect lists source-owned/fail closed.
- Update PPJ schema, capability registry, reference docs, coverage, matrix, and focused regression evidence.

## Capabilities

### New Capabilities

- `ppj-source-bound-text-shadow-blur`: bounded authored and source-bound direct rich-text outer-shadow blur field with token-preserving edit and reprojection.

### Modified Capabilities

- None.

## Impact

- Native PPJ leaf projection, source-bound edit proof/read/patch dispatch, and focused codec tests.
- PPJ schema/capability registry and presentation Skill/coverage documentation.
- No protocol version or host-PowerPoint behavior change.
