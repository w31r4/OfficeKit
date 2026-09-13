## Context

See proposal.md and the source-bound requirements in `specs/presentation-theme-accent1-sat-mod/spec.md`. The native codec already models authored saturation modulation as a bounded thousandth-percent field and already has source-bound accent1 paths for adjacent transforms. The imported theme graph must remain opaque except for an exact direct leaf that can be patched without reserializing unrelated XML.

## Goals / Non-Goals

**Goals:**

- Reuse the existing authored `SaturationModulationThousandth` representation.
- Keep projection, capability authorization, validation, and compilation aligned on one PPJ path.
- Preserve package bytes outside the owning ThemePart and preserve the ThemePart XML outside the direct `satMod/@val` token.

**Non-Goals:**

- Supporting `satOff`, combined transform edits, arbitrary theme XML creation, or host color-management equivalence.
- Changing protobuf contracts or protocol versions.
- Generalizing the profile to accent2 through accent6 in this slice.

## Decisions

- **Use a dedicated source-bound field and capability.** This matches the existing tint, shade, luminance, and alpha transform slices, making authorization field-qualified and preventing accidental broad theme ownership. A generic transform map would make topology checks and capability review less precise.
- **Require one direct RGB leaf and one direct `a:satMod`.** The compiler can then splice one known attribute while preserving all other XML. Accepting transformed colors, nested leaves, or extra siblings would require reconstructing unsupported topology.
- **Map PPJ fractions to integer thousandths with away-from-zero rounding.** `satMod` is non-negative and uses the existing `0..100000` DrawingML scale; the shared authored representation avoids a new wire field.
- **Use the existing ThemePart ownership checks and changed-part accounting.** Only the canonical shared ThemePart is eligible, and no-op/edit tests can prove the rest of the package remains unchanged.

## Risks / Trade-offs

- [Risk] A producer emits a semantically equivalent but structurally different theme color tree. → Mitigation: omit the field and fail closed rather than normalize unsupported topology.
- [Risk] A fractional value rounds at the boundary. → Mitigation: require finite `0..1` input, round deterministically, and reproject the exact integer-derived fraction.
- [Risk] The profile remains intentionally narrow. → Mitigation: document the bounded capability and leave adjacent transform families for separate atomic changes.

## Migration Plan

No data migration is required. Existing authored programs remain valid; imported programs gain the field only when the strict topology matches. Rollback is a normal code revert because the protobuf contract and package format are unchanged.
