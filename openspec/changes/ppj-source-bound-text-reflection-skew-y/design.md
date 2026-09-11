## Context

Direct run reflection projection already has strict full-span proof and independent leaves for position, fade direction, both scale axes, and horizontal skew. The next scalar can reuse that proof without broadening shape, image, chart, or default-text ownership.

## Decision

Add `skewY` to the general reflection value parser for authored PPJ, but opt it into the direct `RunStyle` projector only. Introduce `textReflectionSkewY` with degree scale `60000`; reuse the existing strict skew range and canonical integer validation. The direct source-bound reader accepts `ky` only when the run reflection is otherwise full-span and transform-free, and the edit plan token-splices `reflection/@ky`.

The profile rejects fade, both scales, `kx`, alignment, and `rotWithShape` alongside `ky`, because combining transforms would require a larger proof and could change unsupported effect semantics. It also rejects unknown or duplicate native topology. No deletion or host-rendering behavior is added.

## Lifecycle

1. Authored JSON parses `reflection.skewY` into signed native 1/60000-degree units.
2. Direct run projection emits the PPJ degree and `textReflectionSkewY` only after the strict owner proof.
3. A leaf edit replaces the existing `ky` token in the owning SlidePart.
4. The focused test checks changed parts, Open XML validity, non-target ZIP bytes, and a second projection.

## Risks and boundaries

The profile does not infer missing `ky`, normalize arbitrary effect lists, edit paragraph `defaultText`, or claim PowerPoint visual acceptance. Unsupported graphs remain source-owned.
