## Context

See proposal.md. Native custom paths have optional ExtrusionAllowed mapped to DrawingML extrusionOk; PPJ shared path lowering and projection currently omit it.

## Goals / Non-Goals

Retain the optional path attribute across PPJ lifecycle. This attribute permits native extrusion; it does not author depth, material or 3-D appearance.

## Decisions

Use `extrusionOk` beside fill/stroke to match the XML concept, with no default injection. Reuse native optional presence and existing geometry.paths source authority rather than a new operation. Shared path lowering/projection also retains the attribute where existing mask paths are supported. Keep all topology guards unchanged.

## Risks / Trade-offs

False mistaken for absence -> test all three states. Path rewrite loses unrelated flags -> assert unchanged commands/fill/stroke and non-target ZIP bytes. Preview overclaims 3-D -> assert explicit field diagnostics.

## Migration Plan

Additive PPJ schema; existing native wire needs no regeneration. Refresh projected programs for the new field.
