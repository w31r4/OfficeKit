---
name: "excel-live-control"
description: "Control a workbook already open in Microsoft Excel desktop through OfficeKit's local Excel Add-in. Use for the current open workbook, selected range, or unsaved Excel changes. Use Spreadsheets for standalone .xlsx, .csv, or .tsv files."
---

# Excel Live Control

Use the shared `../office-kit/references/workspace.md` contract for request files
and evidence. Keep JSON requests under `taskRoot`; the live workbook remains
owned by Excel, and screenshots or audits belong under `evidenceRoot`. Return
the session result and evidence without host-specific citation syntax.

## REPL route for an open workbook

For a multi-step live task, launch `officekit repl` and read
`../office-kit/references/repl.md`. Run `await ctx.excel.doctor()` and
`await ctx.excel.sessions()` before a mutation, then pass the same typed JSON
request to `await ctx.excel.execute(request)`. Read the changed range again
after every write; if the response says `maybeApplied: true`, inspect before
retrying. Call `await ctx.excel.disconnect(sessionId)` when the task is done.
This facade does not install the add-in or certificate: keep
`officekit excel install` and `officekit excel uninstall` as explicit control
plane commands.

Use this Skill only for a workbook already open in Microsoft Excel desktop on
Windows or macOS. It operates the open workbook through OfficeKit's local
Add-in and a loopback bridge. It does not use a cloud relay, account, tenant,
or organization deployment.

For a standalone spreadsheet file, use **Spreadsheets** instead. Do not switch
between the two paths unless the user asks to change the target.

## First Connection

Start every task with:

```sh
officekit excel doctor --json
```

If it reports `not-installed`, explain that the local certificate must be
trusted before Excel can load the add-in. With the user's explicit approval,
run:

```sh
officekit excel install --yes --json
```

Then the user completes the one-time Excel step:

1. Open the intended workbook in Microsoft Excel desktop.
2. Choose **Home > Add-ins > My Add-ins > Upload My Add-in**.
3. Upload the manifest path reported by `install`.
4. Open **OfficeKit** from the Home ribbon and choose **Connect OfficeKit**.

The task pane shows the connected workbook, session ID, diagnostics, and a
content-free audit summary. It can be hidden after connection; the shared runtime
keeps the live session active. If the local bridge has stopped while Excel was
closed, run `officekit excel doctor --json` before reopening the OfficeKit pane.

Do not claim that a connection exists until `officekit excel sessions --json`
lists exactly one intended workbook. If multiple workbooks are open and the
target is ambiguous, ask the user which workbook to use.

## Work Loop

1. Discover sessions with `officekit excel sessions --json`.
2. Read the relevant ranges or object metadata before changing anything.
3. Write one local request JSON file using the protocol in
   [references/live-protocol.md](references/live-protocol.md).
4. Run `officekit excel execute request.json --json`.
5. Read the changed range or object again and verify values, formulas,
   formatting, chart/table/PivotTable state, or the screenshot as appropriate.
6. Call the `save` operation only when the user explicitly wants the workbook
   saved. For an unsaved workbook, Excel opens its native Save As prompt and
   the user chooses the file location; OfficeKit never chooses a path or
   overwrites a file on its own.

Use one stable `idempotencyKey` for a single intended operation. Requests are
serialized per session. When an error has `maybeApplied: true`, inspect the
touched target before deciding whether to retry; never assume the operation did
not happen.

For dashboards, charts, dense tables, and substantive visual changes, call
`read_range_image` and review the rendered range. Apply
[style_guidelines.md](style_guidelines.md) and
[charts.md](charts.md) when their subject is relevant.

## Operation Boundaries

The protocol exposes typed primitives for range reads/writes/clears, search,
sheet and workbook metadata, copy, screenshot, dimensions and views, formats,
tables, charts, PivotTables, object lists, and explicit save. Read the
reference before choosing an operation. Do not send JavaScript source code or
unvalidated arbitrary Excel API calls through this route.

An `unsupported-capability` result means the open Excel client cannot provide
that typed operation. Report it plainly; do not simulate the operation with
mouse clicks, macros, COM/VBA, or an unrelated workbook.

The bridge records only timestamp, session, workbook name, operation, range
summary, outcome, and a request hash. It does not record cell values or
formulas. `disconnect <session-id>` ends one pairing. `uninstall --yes`
removes OfficeKit's local state and certificate trust; Excel may still show the
sideloaded item until the user removes it from **My Add-ins**.

## Workbook Quality

- Before editing an existing workbook, read the affected ranges and preserve
  its established formula, format, table, and chart conventions.
- Verify formulas, values, object existence, and key layout after each
  meaningful write. Do not treat a successful command as proof of a correct
  workbook.
- Keep derived values as formulas when the workbook's model calls for formulas.
- For a read-only question, use a read primitive only and make no changes.
- Do not enable macros, bypass protection, alter Excel settings, or claim Web,
  mobile, VBA/COM, or enterprise deployment support.

## Domain Guidance

Read the matching file only when relevant:

- Finance and investment banking:
  [domain_guidance/financial_models.md](domain_guidance/financial_models.md)
- Corporate finance and FP&A:
  [domain_guidance/corporate_finance_fpa.md](domain_guidance/corporate_finance_fpa.md)
- Healthcare: [domain_guidance/healthcare.md](domain_guidance/healthcare.md)
- Marketing and advertising:
  [domain_guidance/marketing_advertising.md](domain_guidance/marketing_advertising.md)
- Scientific research:
  [domain_guidance/scientific_research.md](domain_guidance/scientific_research.md)
