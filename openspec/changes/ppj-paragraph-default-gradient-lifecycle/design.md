## Context

See proposal.md. The gradient schema already owns linear/radial, angle and 2..16 ordered stops. Projection retains stop opacity in native 1/100000 units. The default-run reader/writer accepts direct RGB gradients and canonical centered radial fills; unknown paint remains source-owned. The authored gradient builder currently drops explicit opacity 1 and can round an angle up to a full turn before native validation.

## Goals / Non-Goals

Complete field presence and transitions without rebuilding untouched paragraph paint. Existing literal gradient topology is retained; arbitrary path/tile geometry, imported theme-stop graphs and host text rendering need their own modeling.

## Decisions

Add independent gradient authority and masking beside color. Compare raw gradient values per paragraph, clone the requested gradient only for changed paragraphs, and resolve declared grammar stop colors consistently with the source-bound color field.

Treat color and gradient as two mutually exclusive fields. Switching paint requires deletion of one and assignment of the other, so both authorities are checked. Extend the targeted native fill guard to recognize modeled gradients and simple solid colors; keep rejecting source luminance transforms, duplicate fills and unmodeled paint.

Preserve explicit stop opacity 1 in the authored builder as well as zero, and normalize rounded angles back into one turn. Reuse existing gradient wire fields and native validation.

Use lifecycle fixtures plus a source-decoration/transition case and unsupported source gradient cases. Compare non-target XML/ZIP and fresh projection; retain explicit partial preview diagnostics instead of inferring host rendering from codec tests.

## Risks / Trade-offs

Untouched gradient decorations or alpha precision lost through reserialization → per-paragraph comparison and adjacent native XML evidence.

Known source solid transforms overwritten during a transition → retain the direct solid guard and add transition rejection.

Shared authored gradient normalization affects other owners → run existing gradient regressions alongside the paragraph lifecycle tests.
