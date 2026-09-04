# Design

## Context

DrawingML preset shapes store user-controlled geometry in an ordered
`a:avLst` of named guides. PPJ deliberately exposes only the ordered numeric
values because guide names and formulas are native implementation detail. The
current schema advertises that array, while the compiler rejects it and the
projector omits it.

The normative preset definitions use both percentage-like values and angles in
1/60000-degree units. Existing PPJ's `[-100000, 100000]` range cannot represent
the defaults of several already-listed presets, including stars and circular
arrows.

## Goals / Non-Goals

**Goals:**

- Make the existing ordered PPJ adjustment array executable for every preset
  already listed by PPJ whose native adjustment topology is canonical, for
  both ordinary shapes and recognized picture masks.
- Preserve omission as "use the Office preset defaults".
- Restore recognized imported literal adjustments and permit capability-bound
  value and preset-identity edits without exposing native names or formulas.
- Keep unrelated source content byte-stable at the OPC-part boundary.

**Non-Goals:**

- Raw guide names, formulas, handles, XPath, or arbitrary preset names.
- Formula-valued or partial imported adjustment lists; a picture preset/custom
  topology transition remains a separate no-adjustment custom-mask profile.
- Extending the PPJ preset vocabulary in this change.
- Evaluating or reproducing the complete preset-shape geometry algorithms.

## Decisions

### One canonical profile registry

C# owns one registry keyed by the existing PPJ preset name. Each profile stores
the exact ordered DrawingML adjustment-guide names. Profiles with no adjustable
guides store an empty sequence. The registry is used by authored validation,
native read/write, source editability, projection, and documentation so those
surfaces cannot drift independently.

This is preferred to generating `adj`, `adj1`, ... heuristically because some
standard shapes use semantic names such as `hf` and `vf`.

### Literal integer values only

`geometry.adjustments` becomes an integer array bounded to
`[-21600000, 21600000]`. The bound covers a full signed rotation in DrawingML
angle units and all defaults of the current PPJ preset vocabulary. Exact array
length is either zero (native defaults) or the profile's complete guide count.

The native writer emits only `fmla="val N"`. Imported lists qualify only when
their guide order, names, formulas, attributes, and count are canonical.
Formula graphs remain source-owned.

### Dedicated wire state, declarative PPJ state

`PresentationShape` gains an ordered signed integer field for preset
adjustments. PPJ does not gain methods such as `setCornerRadius`; Agents edit the
typed array and the compiler maps it to native guides.

### Narrow imported capability

Canonical editable preset shapes receive `setGeometry` for exactly
`geometry.adjustments`; recognized picture masks receive one bounded
`setImageMask` capability covering preset identity, complete adjustments, and
the separately proven literal custom-path transition. Source-bound compilation
proves the source hash and changes only the existing geometry owner, replacing
the complete adjustment list when a preset identity changes. Noncanonical
native adjustment lists remain modeled for preview where possible but do not
receive this capability.

## Risks / Trade-offs

- **A broad numeric range can exceed a shape's useful handle range.** The
  standard preset formula pins or interprets the value; OfficeKit guarantees a
  bounded native literal, not that every value produces an attractive shape.
- **A future preset may need values outside one turn.** Such a preset must
  extend the language contract explicitly rather than bypassing validation.
- **Rewriting a canonical target `a:avLst` normalizes that target node.** The
  mutation footprint declares the target shape; all non-target parts and
  elements remain preserved.

## Migration Plan

No successful authored PPJ could previously contain non-empty preset
adjustments because build rejected it. The schema correction therefore has no
valid compiled artifact to migrate. Rollback removes the additive wire field
and restores the fail-closed registry boundary.

## Open Questions

None for this tranche. Adding more preset names belongs to a separate vocabulary
extension backed by the same registry.
