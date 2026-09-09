## Context

See proposal.md. Chart text uses the shared XlsxChartTextStyleCodec; ordinary text and chart frames already use PresentationShadow and PptxShadowCodec. The latter currently compiles in Presentation, while chart styles compile in Shared. Current shadow JSON requires geometry and ordinary mapping inserts defaults, which cannot preserve optional chart attributes.

## Goals / Non-Goals

**Goals:** reuse the existing shadow wire representation and native validation, preserving presence across all chart style owners.

**Non-Goals:** arbitrary effect graphs, inner shadows/glow, host layout and SVG paint acceptance.

## Decisions

- Compile PptxShadowCodec in Shared and add a strict XElement bridge for chart character effect lists. This avoids a second shadow parser/model and keeps root JS imports lazy. Existing shape/picture behavior remains intact.
- Add message field 19 to chart text style, retaining wire version 2. Include shadow in semantic comparison, meaningful-style detection and the global font-only restriction.
- Introduce a chart shadow schema using the ordinary property constraints with only color required. Keep property definitions local because the native schema validator resolves only top-level `$defs` references. Dedicated chart mapping preserves geometry and alpha presence instead of changing ordinary text defaults.
- Resolve declared grammar colors before theme fallback. Preserve theme colors without tint/shade; reject unsupported transforms rather than silently discard them. Explicit opacity overrides color alpha, matching existing shadow behavior.
- Reuse existing style-owner/lifecycle fixtures for line and combo, plus vector precedence and strict native negative tests. Regenerate references and run focused gates.

## Risks / Trade-offs

- Moving a shared codec can expose assembly duplication → update both compile lists and run adjacent shadow tests.
- Unknown XML could be lost by typed parsing → reject unmodeled text, comments, attributes and descendants before accepting a direct effect.
- Source-bound optional values can materialize defaults → exercise color-only, explicit zero/false and deletion against original source bytes with fresh projection.
