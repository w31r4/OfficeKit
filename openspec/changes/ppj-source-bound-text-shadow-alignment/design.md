## Context

The existing PresentationML shadow reader already parses `algn` into the typed shadow model and validates the nine DrawingML rectangle-alignment tokens. Direct rich-text run projection currently rejects alignment along with transforms, and the direct-run edit compiler handles only blur, distance, direction, RGB color, and explicit alpha. The change extends that established source-bound profile by one token while retaining its source-byte and SlidePart ownership rules.

## Goals / Non-Goals

**Goals:**

- Issue one `textShadowAlignment` leaf for a strict direct run outer shadow with an existing valid alignment.
- Re-prove the source token and splice only `outerShdw/@algn`.
- Preserve RGB/theme color, geometry, opacity, siblings outside the profile, and non-target package bytes.
- Cover RGB and theme-color projection, stale/invalid values, and unsupported topology in a focused regression.

**Non-Goals:**

- Adding alignment to paragraph `defaultText`, chart text, shape shadows, or image shadows.
- Creating or deleting an alignment attribute or an entire effect graph.
- Supporting shadow scale, skew, `rotWithShape`, arbitrary transforms, or host PowerPoint rendering acceptance.

## Decisions

1. **Reuse the canonical string token.** `textShadowAlignment` carries the exact DrawingML token rather than a new numeric encoding. This matches existing shape and paragraph alignment leaves and lets the edit plan reuse the bounded token validator.
2. **Keep the direct-run profile strict.** Alignment is admitted alongside the already bounded blur/distance/direction profile, but scale, skew, rotation, siblings, and unknown descendants remain rejected. This makes the new leaf independently provable without making other effect graphs editable by implication.
3. **Use one token splice.** The compiler locates the proven run and outer-shadow owner, checks the exact existing `algn` value, and replaces only that attribute value. It does not serialize the XML subtree, so effect order and unknown bytes remain intact.
4. **Retain shared source proof.** The leaf is added to the closed whitelist, text-owner proof set, readback, and dispatch path. A stale source hash or stale expected alignment fails before any output is written.

## Risks / Trade-offs

- [Unsupported imported graphs] Alignment remains omitted when any unmodeled effect or transform is present → keep the existing fail-closed checks and test each boundary.
- [Host interpretation differs] The token edit proves Open XML and projection behavior, not rendered appearance → document host rendering as unverified.
- [Generated references drift] Schema/registry changes can leave the Skill table stale → regenerate the capability matrix and run the presentation maintainer check.

## Migration Plan

No wire-version or package migration is required. Existing PPJ documents remain valid; the new leaf is issued only by a fresh source projection that meets the strict profile. Rollback is the normal Git revert of the focused commit.
