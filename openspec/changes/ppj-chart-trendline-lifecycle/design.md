## Context

See proposal.md. The PPJ schema already permits zero to sixteen trendlines. The source-bound compiler requires both arrays to exist with matching lengths, PPTX topology checks repeat that restriction, and the shared ChartML patcher only replaces same-index entries.

## Goals / Non-Goals

Goals: complete the list lifecycle on existing editable bar/column and line series, including categorical combo primary and secondary series.

Non-goals: change series/point topology, error-bar presence, workbook formulas, trendline labels or effects, or the separate XLSX adapter contract.

## Decisions

- Keep the existing array contract. Omission and an empty array both remove all native trendlines; no command wrapper, extra IDs or wire fields are necessary.
- Validate all existing native trendline owners before changing the list. Malformed or unsupported owners retain the current opaque/read-only behavior.
- Preserve unchanged same-index nodes. Replace changed nodes, remove surplus nodes, and insert new nodes before error bars and data sources in ChartML order. The array order is authoritative; series identities and data remain fixed.
- Remove trendline-count restrictions only at the PPJ/PPTX boundaries. The shared patcher supports growth/shrinkage, while XLSX keeps its explicit adapter-level topology check.
- Verify native XML and fresh projection with embedded PPJ removed; compare every unrelated ZIP entry and ChartML outside the edited trendline list.

## Risks / Trade-offs

- A malformed original node could otherwise be deleted without inspection → parse all old nodes first.
- A newly inserted node can be placed after data by accident → assert insertion before error bars and data and validate the native series with the Open XML SDK.
- Count changes can accidentally move unrelated chart parts → inspect actual package bytes in the focused experiment.
