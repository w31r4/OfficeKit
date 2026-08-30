## Context

PPJ exposes `waterfall` in the chart-type enum, but `PpjAuthoredPresentationCompiler` rejects it before creating native state. PowerPoint's modern Waterfall chart is represented through the newer ChartEx graph, while OfficeKit's proven editable chart compiler owns the standard DrawingML chart graph. A bounded semantic lowering can provide a truthful editable cumulative bridge without adding a second chart engine or exposing raw ChartEx XML.

## Goals / Non-Goals

**Goals:**

- Give PPJ one finite, explicit waterfall data contract.
- Produce a native editable chart whose bars remain data-backed rather than a group of drawing shapes.
- Preserve exact authored waterfall intent through embedded PPJ recovery.
- Reuse current chart fills, strokes, axes, title, plot surfaces, gap width, frame transforms, and accessibility state.

**Non-Goals:**

- Author or semantically import arbitrary Microsoft ChartEx waterfall graphs.
- Support a cumulative path below zero, mixed series, percent stacking, automatic connector lines, or automatic value labels in the first bounded profile.
- Infer totals from category names or silently repair inconsistent cumulative arithmetic.

## Decisions

1. **Represent point meaning explicitly.** One series owns a `pointRoles` array parallel to `values`; each entry is `delta` or `total`. A total value is an absolute cumulative value, while a delta is added to the running total.
2. **Make the three visual roles explicit.** `chart.style.waterfall` contains required increase, decrease, and total role styles. Each role has a label, a non-image fill, and an optional bounded stroke. This avoids hard-coded red/green/blue conventions and lets the deck design grammar remain authoritative.
3. **Lower to a canonical stacked column chart.** The compiler creates one transparent offset series plus increase, decrease, and total series. Non-applicable points are emitted as missing cached observations, preserving one visible role at each category. The result is a standard editable ChartPart, not a raster or shape approximation.
4. **Keep the arithmetic truthful.** Running totals may not cross below zero in this profile. Every total after the first point must equal the computed running total within numeric tolerance. The compiler fails before output instead of drawing a misleading bridge.
5. **Keep unsupported presentation behavior explicit.** Waterfall uses column orientation and stacked grouping. Legend, explicit stacking, trendlines, error bars, markers, secondary axes, smooth/vary-colors, and automatic data labels are rejected rather than partially applied.
6. **Recover intent from embedded PPJ.** Reimporting an OfficeKit-authored artifact recovers `chartType: "waterfall"`, point roles, and role styles exactly. A package without the embedded program projects the canonical lowered chart as an ordinary stacked column chart; modern or irregular third-party waterfall charts remain source-owned.
7. **Use one existing comprehensive contract.** Add one bridge chart to the canonical PPJ fixture and inspect the native four-series stacked chart plus exact PPJ recovery. Do not create a chart-family matrix.

## Risks / Trade-offs

- **[Semantic drift in lowering]** The native chart has four physical series while PPJ has one semantic series. → Keep stable embedded PPJ authoritative and make ordinary semantic projection describe the native chart honestly.
- **[Misleading negative bridge]** Standard positive stacking cannot represent every path across zero. → Reject any cumulative value below zero in this bounded profile.
- **[Hidden offset leaks into presentation]** A default outline or legend entry could expose the offset series. → Emit explicit no-fill plus zero-opacity outline and reject legends.
- **[Agent assumes ChartEx parity]** The editable result is a standard chart, not a Microsoft modern-waterfall object. → State the lowering in `ppj.md`, chart guidance, and coverage.

## Migration Plan

The schema addition is optional and backward compatible within `office-kit/ppj/v1`. Existing PPJ files and imported PPTX behavior do not change. Rollback is a normal revert; authored waterfall files retain their embedded source program but require a compiler containing this capability to rebuild.

## Open Questions

None for the bounded profile. Connector lines, signed baselines, and selective value labels require separate evidence before expanding the contract.
