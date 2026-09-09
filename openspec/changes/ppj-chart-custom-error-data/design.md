## Context

See proposal.md for the missing PPJ surface. Native error-bar artifacts already have plus/minus values, formats and formula strings, and strict ChartML readers/writers. PPJ currently maps only the four scalar/no-data modes.

## Goals / Non-Goals

Own complete local numeric data for each error side. Workbook/formula synchronization remains an independent ownership problem; formulas will not be relabeled as local arrays. No new wire message or ad-hoc OOXML escape hatch is needed.

## Decisions

- Use `{ values, formatCode? }` for each side rather than bare arrays so the existing native numeric format has a stable owner. Values align with series point indexes; token resolution is used only for the optional string format.
- Add the custom enum and a reusable schema definition, and validate selected sides and array lengths semantically. Native validation remains the final guard for finite numbers and format constraints.
- Reuse native plus/minus artifacts in both compilers and project only formula-free custom owners. The existing unprojected-owner guard then protects formula-backed originals.
- Allow native custom insertion/deletion only when both sides are formula-free. Existing same-presence native formula edits keep their shared codec behavior; the PPJ compiler does not expose them.
- Use the existing lifecycle fixture for four chart profiles and original-source-byte projections. Add focused malformed cache/formula tests and extend the existing preview diagnostic test with custom data.

## Risks / Trade-offs

- Cache lengths or excluded sides are silently coerced → reject inconsistent input before compilation and test negative cases.
- A formula cache is flattened → omit its entire owner from the literal projection and preserve exact ChartPart bytes on no-op.
- Source-bound side changes erase other chart data → compare XML with only the target errBars removed and compare all other ZIP entry bytes.
- Shared codec relaxation affects XLSX → retain adapter presence restrictions and run the shared owner regressions.

## Migration Plan

Additive PPJ schema extension using existing protobuf fields. Older scalar documents retain their behavior. Generated PPJ reference and capability matrix are regenerated with repository scripts; custom data remains explicitly partial in local SVG preview.
