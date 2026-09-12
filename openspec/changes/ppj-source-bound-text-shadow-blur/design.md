## Context

The authored compiler and semantic projector already understand `PresentationShadow`, while native leaf projection and edit-plan dispatch currently cover shape/image and paragraph-default shadow fields only. Direct run import uses the strict no-transform `PptxShadowCodec` profile, which is the intended owner boundary for this increment.

## Goals / Non-Goals

**Goals:**

- Add one direct-run native leaf for `outerShdw/@blurRad` with the existing source-bound proof and token-splice lifecycle.
- Keep the direct run semantic shadow intact and preserve all unrelated XML and ZIP members.
- Add a focused authored → source-bound edit → second projection regression and update registry/docs.

**Non-Goals:**

- Other direct-run shadow fields such as distance, angle, alignment, color, opacity, rotation, scale, or skew.
- Shadow creation/removal, transformed or multi-effect graphs, WordArt, automatic fields, and host-PowerPoint visual acceptance.

## Decisions

1. **Use `textShadowBlurRadiusEmu` as the first direct-run leaf.** It is an existing numeric EMU representation shared by other shadow owners and can be token-spliced without rebuilding the effect list. A complete `shadow` object edit would widen the owner and risk rewriting unrelated native attributes.
2. **Reuse the strict direct outer-shadow topology.** The reader/proof must require one effect list containing one `outerShdw`, one direct RGB/theme color child, no transforms (`sx`, `sy`, `kx`, `ky`), and an existing bounded `blurRad`. This matches current direct-run import behavior; complex graphs remain opaque.
3. **Patch only `blurRad`.** The edit plan reads the existing native value, validates a changed non-negative bounded EMU token, and replaces the original attribute text in the owning `sp` XML. This preserves attribute spelling/order, effect siblings (when the profile admits none), and all non-target package parts.
4. **Keep semantic and native representations aligned.** Authored run styles continue to use the existing points-based `shadow.blur`; projection emits that semantic value from the parsed EMU field while the native leaf carries the exact integer token.

## Risks / Trade-offs

- [Risk] A source file with a valid shadow plus an unmodeled sibling effect can be parsed by another effect reader but cannot safely expose a shadow leaf. → Mitigation: require the strict one-child effect-list proof and omit the leaf for every mixed or malformed graph.
- [Risk] Point-to-EMU conversion can differ at fractional boundaries. → Mitigation: use the existing codec rounding and assert the exact native integer in the focused test.
- [Risk] The test proves package structure and source preservation, not host rendering. → Mitigation: record host acceptance as unverified in coverage/docs.

## Migration Plan

No migration or protocol-version change is required. The field is additive; older readers ignore the new native leaf, while unsupported source topology remains source-owned. Rollback is reverting the single commit.
