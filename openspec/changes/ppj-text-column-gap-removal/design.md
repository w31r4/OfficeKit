## Context

columnGap is a 0..10000 point number, projected from ColumnSpacingEmu. NoColumnSpacing already deletes spcCol and normalizes after edits. PPJ omission and bounded marker admission are missing; tables construct requested bodies separately.

## Goals / Non-Goals

Complete direct numeric spacing presence across existing owners, preserving columns and columnDirection. Full host column layout, inherited geometry and other deletions remain separate work.

## Decisions

Infer the existing marker from old-present/new-absent in both source paths and admit valid true markers. Extend the simple-style whitelist. Reuse the shared lifecycle fixture's compile/project/native/XML helpers for a numeric case, including exact 12700 EMU per point. Setting zero instead of deleting would materialize an override.

## Risks / Trade-offs

Verify zero and fractional values separately from absence, including schema maximum. Keep authority/other-field deletion guards. Table compact restoration must retain native paragraph/run topology. Preserve explicit column count/direction and compare surrounding XML and non-target ZIP bytes.
