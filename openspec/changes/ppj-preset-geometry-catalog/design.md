# Design

## Context

DrawingML defines a finite preset-shape vocabulary. PPJ currently duplicates a
small hand-picked subset in its JSON Schema and in a C# dictionary. A separate
profile file now owns adjustment order and defaults, but it still covers only
that subset. Picture masks reuse the same native `a:prstGeom` construct while
the wire model stores only a preset name, so non-default mask geometry is lost
from the semantic projection.

## Goals / Non-Goals

**Goals:**

- Make every non-connector DrawingML preset in the pinned public schema a
  source-free PPJ shape geometry.
- Preserve the existing `flowChartData` alias while projecting one stable PPJ
  name for each native token.
- Apply the same ordered adjustment contract to picture masks.
- Permit a source-bound mask edit only when the native list is complete,
  literal, canonical, and capability-issued.

**Non-Goals:**

- Inventing shape aliases from UI-localized Office labels.
- Serializing raw guide names, formulas, handles, or arbitrary native tokens.
- Treating connector presets as ordinary shapes.
- Authoring custom-path picture masks.

## Decisions

### The registry owns native identity

Each profile stores its stable PPJ name and native DrawingML token in addition
to ordered guides, defaults, and parameter labels. The codec constructs the
pinned Open XML enum value from that validated token. This removes the second
hand-written C# mapping without introducing runtime reflection.

The schema enum, codec vocabulary, generated reference, and maintainer checks
all derive from the same checked-in registry. The registry is generated once
from the public preset definition resource and then reviewed and versioned in
the repository; ordinary builds do not fetch the network.

### Declarative masks reuse geometry

`image.mask` continues to use PPJ's existing geometry object. For preset masks,
`adjustments` has exactly the same order, units, bounds, defaults, and arity as
the corresponding shape. The additive wire field carries only ordered signed
integers.

### Imported mutation remains narrow

An editable canonical picture receives `setImageMask` only when its preset can
be projected and its `a:avLst` is canonical. Source-bound compile keeps the
mask preset fixed and allows only the complete adjustment array to change.
Custom geometry, partial lists, formulas, unexpected attributes, and topology
changes remain opaque and fail closed.

## Risks / Trade-offs

- The larger vocabulary makes the generated table longer. The short Skill
  continues to route by intent; Agents search the generated reference instead
  of loading every profile eagerly.
- A global numeric safety bound is wider than many individual handles. Office
  preset formulas clamp or interpret the value; the reference publishes native
  defaults and names but does not claim every bounded value is attractive.
- Some preset shapes are rarely useful. They are a finite native vocabulary,
  not a recommendation to decorate pages with them.

## Migration Plan

Existing PPJ names remain valid. The `flowChartData` alias continues to compile
to `flowChartInputOutput`. Newly imported native shapes use a deterministic
preferred PPJ name. Programs with omitted picture-mask adjustments continue to
use native defaults.
