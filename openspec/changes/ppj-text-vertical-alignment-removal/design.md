## Context

PPJ top/middle/bottom maps to native top/center/bottom. NoVerticalAnchor already removes bodyPr anchor and normalizes after edits, but PPJ omission and bounded marker admission are missing. Tables construct bodies separately.

## Goals / Non-Goals

Complete direct verticalAlignment presence across existing owners, preserving anchorCenter. Full host text layout, inheritance and other property deletion remain separate work.

## Decisions

Infer the existing marker from old-present/new-absent in both source paths and admit valid true markers. Extend the simple-style whitelist. Reuse enum lifecycle and table normalization; materializing top instead of deleting would lose native absence. Keep PPJ middle spelling and verify its native Center mapping explicitly.

## Risks / Trade-offs

Whole-style deletion must keep other-field and capability guards. Preserve an explicit anchorCenter in isolated alignment edits. Test native presence, all restored values, source no-op, surrounding XML and non-target ZIP; restore compact table text through identical native paragraph/run topology.
