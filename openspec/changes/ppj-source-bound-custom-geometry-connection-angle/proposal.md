## Why

F-04 already preserves recognized custom-geometry connection sites, but PPJ
does not expose even a safe direct angle field for them. A literal
`a:cxn/@ang` is independently addressable without evaluating the guide graph
or changing connection-site topology, so it is a useful next bounded leaf.

## What Changes

- Expose `customGeometryConnectionSiteAngle60000` for one ordered custom-geometry connection site whose direct `ang` token is a bounded literal integer.
- Token-splice only the selected `a:cxn/@ang` value in the owning SlidePart and re-prove the source binding before writing.
- Preserve connection-site positions, guide references, handles, paths, formulas, list order, and all non-target parts.
- Add a focused authored/imported/source-bound/reprojection regression plus PPJ capability, schema, coverage, and reference documentation.

## Capabilities

### New Capabilities

- `ppj-source-bound-geometry-connection-angle`: One bounded source-bound native leaf for a literal custom-geometry connection-site angle.

### Modified Capabilities

None.

## Impact

The PPJ native-leaf projection and capability registry gain one numeric leaf;
the PPTX edit-plan proof/patch path gains one `a:cxn/@ang` owner. No protobuf
wire-version or general custom-geometry formula/topology behavior changes.
