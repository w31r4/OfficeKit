# Proposal: source-bound shape and image soft edge

Close the next F-05 effect gap with a narrow source-bound `softEdge` owner for
ordinary shapes, lines, and pictures. A direct bounded `a:softEdge` can be
projected as PPJ state and edited through one native radius leaf while the
rest of the SlidePart remains byte-preserved.

This slice intentionally accepts only a direct soft edge, optionally preceded
by one valid outer shadow or one full-span reflection. Other effect graphs,
topology changes, and source-bound insertion/removal/reordering remain
source-owned or fail closed.
