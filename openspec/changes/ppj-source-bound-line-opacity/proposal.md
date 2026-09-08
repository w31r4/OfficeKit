# Proposal: expose source-bound line opacity

Add `lineOpacityThousandthPercent` as an independent PPJ native leaf for the
existing direct line alpha on recognized shape and connector owners.

The existing `line.opacity` semantic field remains the authored and projection
surface. The native leaf makes an already typed value independently editable in
a source-bound package without rebuilding the owner or changing its geometry,
paint, endpoint, or effect topology.
