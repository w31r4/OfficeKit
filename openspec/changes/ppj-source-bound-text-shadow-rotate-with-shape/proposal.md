## Why

Direct rich-text runs can already carry `a:outerShdw/@rotWithShape`, and the PPJ
text shadow object already has the authored boolean. Imported runs currently
drop this explicit flag because the bounded direct-run shadow profile rejects
it, leaving a remaining P0 F-03 text-effect field without a source-bound owner.

## What Changes

- Expose `run.style.shadow.rotateWithShape` through the
  `textShadowRotateWithShape` native leaf.
- Accept only one direct `a:outerShdw` with a canonical `rotWithShape` token,
  required bounded geometry and one RGB/theme color, with no scale/skew
  transform, sibling effect, or unknown descendant.
- Compile a source-bound edit by replacing only the proven
  `outerShdw/@rotWithShape` token in its owning SlidePart.
- Add the focused authored/source-bound/reprojection regression and update the
  schema, capability registry, generated references, coverage, backlog, and
  presentation Skill docs.

## Capabilities

### New Capabilities

- `ppj-source-bound-text-shadow-rotate-with-shape`: Direct imported rich-text
  run outer-shadow rotation flag with a bounded source-preserving edit.

### Modified Capabilities

- None.

## Impact

The direct-run shadow safety predicate, projection, native-leaf proof/read and
XML patch dispatch, PPJ native-leaf schema/registry, generated capability
matrix, presentation references, coverage/backlog notes, and one codec test
are affected. The existing wire field and authored compiler are reused; no
protocol version change or host-PowerPoint rendering claim is required.
