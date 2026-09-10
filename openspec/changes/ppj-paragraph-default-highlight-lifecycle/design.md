## Context

See proposal.md. The authored color resolver already produces opaque RGB highlights. Native default styles carry either RGB or scheme highlight, while the current projector emits only RGB. A shared TryHighlight reader recognizes simple color nodes and withholds transformed/unknown graphs.

## Goals / Non-Goals

Complete direct RGB and theme highlight presence without rewriting other paragraphs, direct runs or source content. Imported transforms, alpha-bearing highlight graphs and host glyph rendering keep their own boundaries.

## Decisions

Extend exact field authority and masks. Compare raw highlight presence/value per paragraph before changing its cloned style: rebuilding all paragraphs from authored resolution would convert untouched scheme values to RGB.

Project native scheme highlight as a token object. On a changed source-bound highlight, preserve an untransformed standard theme token unless an explicit grammar color token shadows it; other accepted colors use existing opaque RGB resolution. Reject nonopaque results.

Extract and reuse targeted highlight writing in the scalar and whole-style writers, retaining the existing native order and unknown-graph guard. Tighten the shared topology reader for hidden non-element content instead of allowing it to vanish.

## Risks / Trade-offs

Theme conversion in untouched paragraphs → one source fixture edits a neighboring RGB paragraph and compares XML/ZIP.
Theme/grammar token ambiguity → test standard theme, declared token override and tint/shade resolution separately.
Unknown content erased during removal → source fixtures check no-op identity, rejected replacement and unrelated scalar edits.
Whole-default-style baseline failure → retain the documented exclusion; do not claim full host acceptance.


Source experiment boundary: the SDK preserves unknown child elements but drops
illegal bare character data inside a DrawingML color node during deep parsing.
Keep source binding rejection for that malformed text case; do not normalize its
hash or emit a lossy candidate. A separate unknown-element fixture verifies
preservation during unrelated edits.
