## Context

See proposal.md. from_word_art uses optional field 36 and native fromWordArt. Text warp is modeled independently through textWarpPreset and adjustments.

## Goals / Non-Goals

Complete marker presence while preserving independent text warp. Full WordArt geometry/effects/host rendering remain separate.

## Decisions

Add no_from_word_art at unused field 44, retaining setter 36. Reject false deletion and concurrent setters. Apply and normalize deletion through the optional-boolean pattern. Reuse shared lifecycle/wire tests with textArchUp retained; move remaining whole-style rejection to flatTextZ.

## Risks / Trade-offs

Deleting a marker could wrongly remove warp → include independent warp in source fixtures and compare native XML. Old codecs ignore new fields → document updated codec requirement. Concurrent preview registry edits → stage only owned reasons.
