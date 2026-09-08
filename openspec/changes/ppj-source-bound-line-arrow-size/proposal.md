# Proposal: expose bounded line arrow size leaves

Add source-bound native leaves for the explicit DrawingML width and length
tokens on a recognized line head or tail. Keep the existing PPJ arrow type
surface and wire compatibility unchanged; the new leaves are independently
editable capability fields for imported shape and connector owners.

The profile is deliberately token-local. It does not infer an arrow from
geometry, change arrow type, or rebuild a line style. Missing, inherited,
malformed, extension-bearing, or otherwise ambiguous arrow endpoints remain
source-owned.
