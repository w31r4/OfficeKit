## Context

See proposal.md. compatible_line_spacing uses optional field 35 and native compatLnSpc. Optional boolean deletion has established source-inference, validation and normalization paths.

## Goals / Non-Goals

Complete direct hint presence while retaining actual paragraph line spacing. Host line metrics remain separate.

## Decisions

Add no_compatible_line_spacing at unused field 43, preserving setter 35. Reject selected false and concurrent setters. Apply deletion to compatLnSpc and normalize intent before semantic comparison. Extend shared boolean/wire tests with explicit line spacing; move the remaining whole-style rejection fixture to fromWordArt.

## Risks / Trade-offs

Hint deletion might change paragraph spacing → compare actual native paragraph XML. Old codecs ignore new fields → require updated codec. Concurrent native build → wait for its live process to terminate before editing source. Shared registry changes → stage only owned reasons.
