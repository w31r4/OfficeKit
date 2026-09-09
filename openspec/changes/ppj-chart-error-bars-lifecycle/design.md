## Context

See proposal.md for the F-07 gap. The source-bound PPJ compiler, PPTX topology comparison and shared native patcher each currently require unchanged error-bar presence. Custom native plus/minus data can be read but is not projected into this PPJ object.

## Goals / Non-Goals

Enable the existing four-mode object lifecycle without new wire fields or a new operation. Custom-data authoring and workbook/formula synchronization remain separate backlog work; XLSX adapter presence policy stays unchanged.

## Decisions

- Treat object omission as deletion, consistent with other optional PPJ fields. Reuse the existing analytics capability and scalar parser instead of introducing a separate delete command.
- Remove only the error-bar presence comparison in PPTX topology checks. Keep series counts, data lengths, missing points, plot families and axis groups protected.
- Validate any existing native owner before removal or replacement. Insert a missing scalar owner immediately before category/value data, after trendlines, preserving the native series order. Reject duplicated owners and custom presence changes in the shared patcher.
- Check the original native errorBars in the PPJ compiler before assignment. An unprojected custom owner must not be mistaken for an empty slot; rejecting this edit preserves its plus/minus caches and formulas.
- Reuse the existing two-series chart fixture. Test column, horizontal bar, line and secondary-axis combo through native compile and fresh projection from original package bytes with the private PPJ snapshot removed.
- Keep the SVG preview boundary explicit: an errorBars owner produces a partial diagnostic until its geometry is rendered. A dependency-light preview regression checks presence and removal; it does not stand in for codec or visual validation.

## Risks / Trade-offs

- Shared patcher behavior affects two formats → retain the existing XLSX adapter guard and run shared owner regressions.
- Invalid original owner erased during deletion → validate before any mutation and verify unchanged native XML on rejection.
- Field tests overstate completion → report structural and projection evidence only; keep custom data and remaining F-07 gaps visible.

## Migration Plan

No schema or package migration. Existing programs remain valid. Reverting the implementation restores the previous presence restriction without altering existing file data.
