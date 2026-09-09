## Context

See proposal.md. Existing cache/workbook token splicing is limited to opaque chart data leaves. The typed PPJ chart path rebuilds its owned ChartPart and has no embedded-workbook mutation yet.

## Goals / Non-Goals

Synchronize error amounts inside existing formula bindings. Arbitrary formula evaluation, reference retargeting and worksheet topology changes remain distinct work. Existing literal error-bar lifecycle stays intact.

## Decisions

- Project complete local formula owners into the existing plus/minus data objects using the existing formula string syntax. Retain the formula exactly; source-bound comparison rejects binding topology changes.
- Plan workbook changes inside PPTX chart export before writing either part. Reuse the local range parser and the existing worksheet XML value-splice writer instead of exporting a reconstructed workbook.
- Require unique ChartPart/workbook relationships and numeric cells equal to the original cache. Reject overlapping chart range consumers and workbook formula/dependency graphs that cannot be kept consistent by numeric cell replacement.
- Rebuild only owned error-bar ChartML and splice only selected worksheet values. Declare the embedded package's changed path and new hash to the existing source-preservation validator, including grouped charts.
- Keep no-op exports byte-identical and style-only error-bar edits ChartPart-only. Use one ordinary and one combo workbook fixture for the positive chain plus focused closure rejection cases.

## Risks / Trade-offs

- Changing a cell affects another consumer → check chart range overlap, shared package relationships and workbook dependency metadata before mutation.
- Partial output on late failure → prepare replacements in memory, publish only after validation, and return no file on compilation failure.
- New source-preserved part rejected by export audit → explicitly include the workbook path/hash in the chart replacement result and test exact outer/inner footprints.
- Overstated completion → report fixed-binding synchronization only; retain the separate formula topology and host/render gaps.

## Migration Plan

Additive PPJ schema using existing wire fields. No external service, runtime download or package rebuild is required for the focused managed experiment.
