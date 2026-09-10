## Context

The native reflection message already stores optional `stPos` and `endPos`.
Direct rich-text runs currently use the full-span reader and issue only the
five existing reflection scalar leaves. Paragraph `defaultText` already has a
strict one-variable-endpoint profile that can be reused for the run owner.

## Goals / Non-Goals

**Goals:**

- Read a direct run reflection with an explicit `stPos` and absent or full-span
  `endPos` as a bounded semantic reflection.
- Issue one run/text-index-bound numeric leaf and splice only `@stPos` after
  proving the source token.
- Keep existing run reflection scalar leaves strict for full-span reflections.

**Non-Goals:**

- Do not add protobuf fields, change authored reflection syntax, or change the
  wire version.
- Do not expose variable `endPos`, two-variable ramps, transforms, WordArt, or
  host PowerPoint rendering as editable capabilities.

## Decisions

Use a dedicated `textReflectionStartPosition` leaf rather than treating the
position as opacity. Its binding stores the canonical `0..100000` token while
the PPJ value remains the existing `0..1` reflection position.

The direct run reader will opt into variable positions only for projection and
then accept the narrow start-position profile: `endPos` must be absent or
`100000`. For a variable start-position reflection, existing blur/opacity/
distance/angle leaves are withheld; this prevents an old full-span proof path
from accepting a graph it cannot safely patch. Full-span runs retain all five
existing leaves.

The source-bound writer will use the existing text-leaf ordinal and direct
`rPr/effectLst/reflection` range, replacing only the `stPos` value. An explicit
full-span `stPos=0` also receives the dedicated leaf so the source token is
represented losslessly; the existing five scalar leaves remain available for
that full-span profile. A malformed,
noncanonical, stale, duplicate, or two-variable reflection stays source-owned.
The focused regression authors a direct run reflection with `stPos=20000` and
`endPos=100000`, removes the embedded PPJ, edits the leaf to `65000`, checks a
SlidePart-only change and valid XML, and verifies a second projection. A
variable-end fixture proves the capability is not widened.

## Risks / Trade-offs

- [Risk] Opting the run reader into variable positions could expose an
  unsupported two-variable graph. → [Mitigation] Filter the run profile to an
  explicit start endpoint with absent/full-span end and withhold all other
  scalar leaves for that graph.
- [Risk] A semantically equal decimal could hide a noncanonical raw token. →
  [Mitigation] Validate the exact integer token and compare it during proof
  before patching.
