## Context

See proposal.md. CustomAdjustments and CustomGuides use the same native guide message; the native validator evaluates adjustments first, then guides, with a shared namespace. PPJ's shape parser already gates preset integer parsing by kind; diagram validation needs the same guard.

## Goals / Non-Goals

Goals: expose ordered custom adjustment formulas for the current literal-path shape profile and complete its adjustment-to-guide-to-rectangle dependency chain. Non-goals: handle/site authoring, reference-backed path coordinates, masks/clips and host interaction.

## Decisions

- Custom geometry.adjustments is a max-256 list of name/formula objects; preset geometry.adjustments remains the existing integer vector. Native formulas retain their own units and validation.
- Empty/omitted lists project as omission. Custom adjustments precede ordinary guides and share duplicate/reference checks; adjustments cannot refer to later guides.
- Extend existing custom setGeometry authority with geometry.adjustments. Source changes rebuild and validate the final graph; removal of an adjustment with remaining dependents rejects, while coordinated cleanup is allowed.
- Permit recognized CustomAdjustments in projection/edit classification while retaining handle/site and nonliteral-path guards. Actual shape owners accept the field; shared mask/clip lowering rejects it.
- Test a two-adjustment chain feeding a guide and text rectangle, actual evaluated change, list addition/removal and cross-list failures. Existing preset regressions verify the kind boundary.

## Risks / Trade-offs

- Preset and custom values confused -> discriminate by kind in schema, parser and diagram validator.
- Cross-list dangling reference -> native final-graph validation.
- Empty state becomes a default guide -> omit empty lists on fresh projection.

## Migration Plan

Additive custom schema and issued field authority with no wire change. Refresh projections and synchronize generated references.
