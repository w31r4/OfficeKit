# Spreadsheets

Spreadsheets is the file-type wrapper plugin for spreadsheet artifact workflows, Google Sheets-ready workbooks, and Microsoft Excel live control.

This installable Skill bundle is distributed with `office-kit`.

## Included Skills

- `Spreadsheets`: create, edit, analyze, visualize, render, and export spreadsheet files such as `.xlsx`, `.xls`, `.csv`, and `.tsv`, including Google Sheets-targeted workbooks that should be authored locally before import.
- `excel-live-control`: inspect, edit, verify, and control a workbook already open in Microsoft Excel desktop through the local OfficeKit Add-in.

## Discoverability

Use this plugin for spreadsheet-oriented terms from the file-type naming model: sheet, sheets, Google Sheets, Excel, CSV, model, spreadsheet, spreadsheets, workbook, tracker, and `.xlsx`.

## Excel Live Control Boundary

The local `Spreadsheets` Skill works on durable spreadsheet files. `excel-live-control` is OfficeKit's optional local-host adapter for an already-open workbook: `officekit excel install` creates a user-local HTTPS certificate and sideload manifest, then the OfficeKit ribbon connects the workbook to the local CLI. It has no account, tenant, cloud relay, or organization deployment requirement. The first runtime load requires access to Microsoft's Office.js CDN.

Excel Live Control V1 supports Microsoft Excel desktop on Windows and macOS. It exposes typed range, formatting, chart, table, PivotTable, screenshot, and explicit-save primitives; it does not provide arbitrary Office.js code execution. The shipped automated bridge suite is supplemented by platform-specific manual Excel acceptance before this surface can be marked fully complete.

The core reference-style workbook example is tested end to end. Full parity with every API named in `artifact_tool_docs/API_QUICK_START.md` is still in progress and is tracked in `docs/reference-skills.md` in the source repository.

## Source

The plugin tree is versioned directly under `skills/spreadsheets` in the public repository.
