## Context

See proposal.md. Chart text already has optional shadow geometry and transforms. Paragraph defaults use the required ordinary profile, a strict shadow-only reader, and lack shadow field authority. The direct-effect patch helper and serialized presence comparison now support the preceding inner-shadow/reflection increments.

## Goals / Non-Goals

Expose all modeled direct shadow values with independent removal/restoration while preserving source XML. Keep ordinary run/shape/image syntax and native-leaf proof profiles unchanged. Host appearance and inherited defaults remain separate work.

## Decisions

Reuse chartTextShadow in the paragraph-default schema and its optional builder/projector. Add the same paragraph color handling as glow/innerShadow: explicit opaque RGBA retains alpha, transformed colors resolve to RGB, and declared source grammar colors take precedence over theme names. This avoids duplicating a partial shadow vocabulary.

Opt the isolated paragraph shadow reader into transforms, with PPJ blur/distance bounds. Retain existing checks for duplicate lists/effects, DAGs, unknown attributes, color graphs and descendants. Allow the existing default-shadow distance leaf to represent the 1,270,000,000 EMU endpoint without broadening unrelated leaf ranges.

Compare serialized shadow values to retain optional zero/false presence. Build the new shadow in an isolated effect list and splice only its node through the shared effect helper; the general shadow writer must never remove a mixed source list. Deletion removes only recognized shadow state and an empty attribute-free wrapper.

## Risks / Trade-offs

Theme/grammar collisions or explicit alpha collapse → exercise theme identity, source grammar precedence and omitted/zero/one alpha through native XML and fresh PPJ.

Transform or geometry bounds hide source-owned state → test invalid native attributes, unsupported color descendants and sibling preservation during unrelated scalar edits.

Mixed effects or direct runs get rewritten → compare source XML and non-target ZIP entries in the existing lifecycle fixtures.

Preview diagnostics are confused with host evidence → retain partial grading and record unperformed host/AOT checks.
