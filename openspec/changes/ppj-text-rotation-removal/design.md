## Context

See proposal.md. Upright already compares old/new styles and emits a deletion marker. Text rotation still merges only present values. Tables have their own style builder; compact-text normalization and fixed-topology restoration now exist.

## Goals / Non-Goals

Preserve rotation absence independently of zero in every currently editable text-body owner. Keep element frame rotation separate, and retain other style deletion and unsupported-owner boundaries.

## Decisions

Use existing NoRotation when previous style contains rotation and the requested style omits it. Admit that valid operation in the bounded native body profile. Extend removable whole-style fields to upright and rotation only; removing a style containing both clears both. Retain existing degree-to-60000 conversion for explicit values and shared table normalization. Extend the existing lifecycle fixture to verify rotation without another rendering or acceptance framework.

## Risks / Trade-offs

Zero collapsed into absence -> assert native rot=0 and fresh projection. Wrong frame owner -> compare all XML except bodyPr rot. Table marker rejected or lost -> original-source table deletion/restoration test. Other fields silently deleted -> retain whole-style whitelist and existing rejection regression.
