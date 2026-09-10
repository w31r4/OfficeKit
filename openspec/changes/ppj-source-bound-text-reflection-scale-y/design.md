## Context

Direct run reflection projection already has strict full-span proof and independent leaves for position, fade direction, and horizontal scale. The next scalar can reuse that proof without broadening shape, image, chart, or default-text ownership.

## Decision

Add `scaleY` to the general reflection value parser for authored PPJ, but opt it into the direct `RunStyle` projector only. Introduce `textReflectionScaleY` with ratio scale `100000`; keep the existing signed-int range and nearest-even rounding. The direct source-bound reader accepts `sy` only when the run reflection is otherwise full-span and transform-free, and the edit plan token-splices `reflection/@sy`.

The profile rejects fade, `sx`, skew, alignment, and `rotWithShape` alongside `sy`, because combining transforms would require a larger proof and could change unsupported effect semantics. It also rejects unknown or duplicate native topology. No deletion or host-rendering behavior is added.

## Lifecycle

1. Authored JSON parses `reflection.scaleY` into the native signed 1/100000 value.
2. Direct run projection emits the PPJ ratio and `textReflectionScaleY` only after the strict owner proof.
3. A leaf edit replaces the existing `sy` token in the owning SlidePart.
4. The focused test checks changed parts, Open XML validity, non-target ZIP bytes, and a second projection.

## Risks and boundaries

The profile does not infer missing `sy`, normalize arbitrary effect lists, edit paragraph `defaultText`, or claim PowerPoint visual acceptance. Unsupported graphs remain opaque.
