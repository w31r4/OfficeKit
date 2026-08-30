## 1. Contract

- [x] 1.1 Add the shared additive wire-v2 image-paint message and owner fields for shape, background, and picture tile mode.
- [x] 1.2 Add `setBackground` and `setImageFit` PPJ capability vocabulary and keep schema, semantic validation, and generated bindings synchronized.

## 2. Native DrawingML

- [x] 2.1 Implement one bounded shape/background image-paint codec for embedded asset, signed source rectangle, direct alpha, stretch, and parameter-free tile.
- [x] 2.2 Extend the picture codec with the same parameter-free tile profile while preserving crop, mask, border, shadow, accessibility, and relationship cleanup.
- [x] 2.3 Extend shape and background read/build/apply/scrub paths without absorbing unsupported blip effects or external links.

## 3. PPJ Compilation And Projection

- [x] 3.1 Lower authored image fill fit/crop/opacity into shape, background, and picture native state using declared asset dimensions.
- [x] 3.2 Project recognized imported image paint into typed PPJ and issue owner-local capabilities.
- [x] 3.3 Lower capability-issued source-bound shape fill, page background, and picture fit changes with asset and source-revision proof.

## 4. Agent Surface

- [x] 4.1 Update the central capability registry and regenerate `ppj.md`.
- [x] 4.2 Update focused Shapes and Media/Layers guidance with executable patterns and fail-closed boundaries.

## 5. Lean Verification And Delivery

- [x] 5.1 Extend one existing integrated PPJ codec test across authored build, reimport, source-bound edit, and second projection.
- [x] 5.2 Run the narrow codec test, proto check, Skill maintainer check, OpenSpec strict validation, and diff check.
- [x] 5.3 Commit by contract/native/compiler/docs/evidence boundaries, push normally, and fast-forward remote main after verifying ancestry.
