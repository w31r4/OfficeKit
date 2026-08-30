## Context

DrawingML text run and default-run properties own one text-fill choice and an optional effect list. OfficeKit currently models only direct solid text fill in those locations, even though `PptxGradientFillCodec` and `PptxShadowCodec` already prove a bounded direct-RGB gradient and outer-shadow profile for other presentation objects.

## Goals / Non-Goals

**Goals:**

- Expose linear and centered-radial gradient text with two through sixteen ordered direct-RGB stops.
- Expose one direct outer shadow with explicit color, opacity, blur, distance and angle.
- Apply the state to direct runs and paragraph default text styles.
- Preserve unknown text fills and effect graphs during unrelated source-bound edits.

**Non-Goals:**

- Theme-color gradient stops, reflection, glow, soft edges, inner shadows, 3D text, WordArt transforms or arbitrary DrawingML effects.
- Inferring a simplified effect from an irregular third-party graph.
- Adding a second text styling language or raw OOXML escape hatch.

## Decisions

1. **Reuse existing PPJ value shapes.** `textStyle.gradient` uses the same bounded gradient object as other PPJ fills, without a redundant `type: "gradient"` discriminator. `textStyle.shadow` reuses the existing shadow definition.
2. **Keep one text-paint choice.** A text style may declare `color` or `gradient`, never both. Color opacity remains part of the color value; gradient opacity belongs to individual stops.
3. **Share the proven native codecs.** The run and default-run readers/writers call `PptxGradientFillCodec` and `PptxShadowCodec`; no independent effect implementation is introduced.
4. **Fail closed around source topology.** A source-bound edit may replace or remove a text fill/effect only when the existing graph is canonical solid, canonical gradient, canonical outer shadow, or absent. Unknown graphs remain source-owned and block only the requested conflicting mutation.
5. **Recover canonical state.** Ordinary PPTX projection emits gradient and shadow only when their native graph is proven. OfficeKit-authored PPTX continues to recover the exact embedded PPJ.
6. **Use one existing contract.** Extend the comprehensive PPJ authored fixture with one gradient/shadow title and a solid-color conflict assertion. Do not add an effect matrix.

## Risks / Trade-offs

- **Text effects can reduce legibility.** The Skill treats them as display-text tools and requires rendered contrast review; body copy remains solid by default.
- **Native child ordering is strict.** The codecs insert fill and effect nodes using DrawingML schema ordering rather than arbitrary append order.
- **Imported effects are broader than the bounded profile.** Unsupported graphs remain opaque; capability claims stay narrower than PowerPoint's full effect UI.

## Migration Plan

The fields are optional and backward compatible in `office-kit/ppj/v1`. Existing PPJ programs compile unchanged. Rollback is a normal revert; programs using the new fields require a compiler that advertises this capability.

## Open Questions

Glow, reflection and WordArt transforms remain separate evidence-driven slices.
