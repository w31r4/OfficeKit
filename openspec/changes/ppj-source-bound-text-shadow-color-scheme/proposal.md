## Why

Direct rich-text runs already preserve theme-colored outer shadows in the
semantic shadow model, but imported source-bound projection only issues the
RGB color leaf. The remaining F-03 sub-gap prevents callers from editing a
direct `schemeClr` shadow token without rebuilding the effect list.

## What Changes

- Expose `run.style.shadow.color` through the `textShadowColorScheme` native leaf.
- Accept only one strict direct `a:rPr/a:effectLst/a:outerShdw` with one bare,
  canonical `schemeClr` token, bounded geometry, and no transforms, sibling
  effects, or unknown descendants.
- Compile a source-bound edit by replacing only the proven
  `outerShdw/schemeClr/@val` token in its owning SlidePart.
- Add the focused authored/source-bound/reprojection regression and update the
  schema, capability registry, generated references, coverage, backlog, and
  presentation Skill docs.

## Capabilities

### New Capabilities

- `ppj-source-bound-text-shadow-color-scheme`: Direct imported rich-text
  outer-shadow theme color with a bounded source-preserving token edit.

### Modified Capabilities

- None.

## Impact

The native text leaf projection and source-bound proof/read/patch dispatch gain
one theme-color sibling. PPJ schema, capability metadata, generated matrix,
presentation references, coverage/backlog notes, and one codec test are
updated. The existing wire field and authored compiler are reused; no protocol
version or host-PowerPoint rendering claim changes.
