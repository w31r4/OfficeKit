## Context

Direct run reflection already has strict full-span proof and independent leaves for positions, fade direction, and scale. The native reflection codec already understands signed `kx`; the missing piece is the narrow direct-run PPJ field and source-bound token owner.

## Decision

Add `skewX` to the general reflection parser and direct-run projector. Introduce `textReflectionSkewX` with degree scale `60000`; reuse the existing strict skew range and canonical integer validation. The direct source-bound reader accepts `kx` only when the run reflection is full-span and otherwise transform-free, and the edit plan token-splices `reflection/@kx`.

The profile rejects fade, both scales, `skewY`, alignment, and `rotWithShape` alongside `kx`. Unknown or duplicate topology remains opaque. No paragraph default-text, shape/image, deletion, or host-rendering behavior is added.

## Lifecycle

1. Authored JSON parses `reflection.skewX` into signed native 1/60000-degree units.
2. Direct-run projection emits the PPJ degree and `textReflectionSkewX` after strict owner proof.
3. A leaf edit replaces the existing `kx` token in the owning SlidePart.
4. The focused test checks changed parts, Open XML validity, non-target ZIP bytes, and a second projection.

## Risks and boundaries

The profile does not infer missing `kx`, normalize arbitrary effect lists, or broaden transform combinations. Unsupported graphs remain source-owned.
