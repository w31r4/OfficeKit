# Presentation design review

Design review separates machine-provable plan invariants from signals that
still require judgment.

## Blocking checks

With `authoringPlan`, OfficeKit can fail review for:

- invalid plan or plan/candidate revision mismatch;
- page-count mismatch;
- required unresolved decisions;
- colors outside a strict declared palette;
- fonts outside a strict declared font set;
- text runs below the declared `minimumCaptionFontSize`;
- page text/object budgets exceeded;
- changed or missing pages outside `changedPageIds`.

Fix these before commit. A new plan revision requires a new reviewed artifact
commit before publication.

## Review warnings

OfficeKit may report:

- repeated modeled composition signatures;
- sharp density changes between adjacent pages;
- six or more same-sized boxes on one page;
- repeated title openings;
- new modeled color or font tokens during a local edit;
- a missing or unfinished visual carrier;
- motion added before the declared composition is complete;
- repeated motion rhythm or excessive pulse emphasis.

These are bounded signals. Inspect the relevant pages and decide whether the
pattern is intentional. If it is intentional, record an explicit
`designGrammar.intentionalWarnings` entry with the warning `type`, affected
`pageIds`, and a concrete `reason` before commit. A prose note elsewhere in the
plan is not enough to make the exception auditable. Do not describe a
warning-free report as proof of good design.

## Visual evidence

When image understanding is available, inspect rendered pages at readable
scale and review hierarchy, crop, balance, contrast, visual rhythm, and the
relationship between copy and graphics.

When it is unavailable, retain semantic, structural, layout, and design checks
and report `visualReview: "unavailable"`. Request human review for high-risk
design decisions.

AnyDoc is a text/table reading view. Use it for a declared content-coverage gap,
such as unavailable visual review or truncated multi-page inspection. It does
not verify pixel layout, images, formulas, animations, or aesthetics.

For animated decks, record `playbackEvidence` separately. `structural` proves
only that the timing graph, targets, builds, and Morph pairs are coherent.
`keynote` or `powerpoint` means that host was actually used for playback. Static
render review remains required because correct timing cannot repair a weak or
broken page composition.

## Local edit evidence

Pass the latest reviewed PPTX as `baseline` and the exact plan page IDs as
`changedPageIds`. Review affected pages visually and verify non-target page
signatures remain stable. If the user asks for a global redesign, omit the local
scope and update the plan first.

For a self-directed or newly composed deck, do not use an earlier draft as a
baseline. Run the final semantic and layout review against the candidate alone
and repair errors such as text overflow or unexpected overlap before commit.
Baseline review may downgrade only an unchanged, already evidenced finding in
a source-bound local edit; it is not a way to make a new deck pass.

Intentional layering is allowed when a filled, text-free rectangle, rounded
rectangle, or ellipse is a bounded background container for its child objects,
or a thin track behind centered markers. Keep those children within the
container or track span. An object that crosses the container boundary remains
an overlap error and must be repositioned.
