# Presentation template restoration index

Image-only sources do not have an original PPTX to diff. OfficeKit therefore
creates an unrelated, editable reference deck and records two separate
weighted indices before a style can be called restored.

## Visual index

The visual score is a weighted review of silhouette (16), hierarchy (14),
palette and surfaces (12), typography (10), density and rhythm (10), visual
carriers (12), layer relationships (12), motifs (8), and example coverage (6).
It compares the authored deck with the source images and written direction. A
high score means the same design decisions are recoverable; it does not claim
pixel identity when the source has no editable package.

Each candidate manifest records SHA-256 values for the source montage and
design guide. Those source files remain outside the package; the hashes make
the clean-room comparison reproducible without treating the source as a
distributable asset.

## Functional index

The functional score weights inspect discovery (14), editable leaves (14),
reusable assets (12), round-trip stability (16), native rendering (14),
background/layer fidelity (12), opaque preservation (10), and safe refusal (8).
The evidence must include a real OfficeKit reference package, inspect output, a
local edit, a second import, and a native render check.

Both indices must be at least 95. A missing source, render, inspect, edit, or
re-import proof is a failed restoration, regardless of the numeric claims. The
source remains descriptive evidence only; it is never copied into a published
Skill.
