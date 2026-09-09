## Context

See proposal.md. The writer replaces literal adjustment lists when a preset setter is present, but source-bound merge retains old values when PPJ fields disappear. The table path already builds a replacement body-style value.

## Goals / Non-Goals

Goals: complete presence lifecycle for the existing preset/literal-guide profile.
Non-goals: arbitrary formulas, full WordArt rendering, inherited formatting resolution.

## Decisions

- Add true-only no_text_warp_preset field 46, preserving setter 37 and repeated guides 38. Reject a simultaneous setter or nonempty guide list.
- Detect source preset omission and clear preset/guides during edit-intent normalization. Omitted or empty guides with a retained preset replace the list with no guides; no separate wire marker is needed.
- Allow empty PPJ adjustment arrays and normalize them to absence after projection.
- Keep canonical source topology checks. Inspect guide XML before typed access because ShapeGuide is an SDK leaf that can hide illegal nested content.
- Extend shared tests for five owners and simple shape/table styles; include dependent-field rejection, ordered/signed/zero guides, restoration and non-target source preservation.
- Once warp fields are removable, all accepted direct body-style fields support whole-style removal. Retire the obsolete test that requires warp whole-style removal to fail, keeping authority rejection.

- The nativeLeaf schema still limits textBodyWarpAdjustment to -1e9..1e9 despite the modeled signed 32-bit guide contract. Extend the existing kind-specific signed 32-bit branch and cover both limits in the shared restoration fixture; other leaf ranges stay unchanged.

## Risks / Trade-offs

Updated codec required. An omitted guide list intentionally removes old explicit guides while retaining the preset. Unknown native markup remains source-owned and editing it must fail closed.
