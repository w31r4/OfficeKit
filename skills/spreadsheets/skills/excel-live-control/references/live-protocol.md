# OfficeKit Excel Live protocol

`officekit excel execute` accepts one local JSON file. It rejects URLs, stdin,
symlinks, and malformed or oversized requests.

```json
{
  "protocol": 1,
  "sessionId": "session-from-officekit-excel-sessions",
  "idempotencyKey": "one-stable-key-per-intended-operation",
  "operation": "write_range",
  "args": {
    "sheet": "Forecast",
    "range": "B4:C4",
    "values": [[125000, 138000]]
  }
}
```

The result is either:

```json
{ "protocol": 1, "ok": true, "result": {}, "audit": {} }
```

or:

```json
{
  "protocol": 1,
  "ok": false,
  "error": {
    "code": "session-unavailable",
    "message": "...",
    "retryable": true,
    "maybeApplied": false
  },
  "audit": {}
}
```

`maybeApplied: true` is a signal to re-read the target before retrying. An
idempotency key returns the original completed result for the same session and
operation window.

## Typed operations

| Operation | Required arguments | Use |
| --- | --- | --- |
| `read_ranges` | `sheet`, `ranges` | Read bounded values, formulas, displayed text, or number formats. |
| `search_workbook` | `query` | Literal search over bounded used ranges; optional `options` supports case, complete-match, and result-count controls. |
| `list_items` | `kind` | List `worksheets`, `tables`, `charts`, `pivotTables`, or `names`. |
| `write_range` | `sheet`, `range`, `values` and/or `formulas` | Write rectangular matrices; `numberFormat` must have matching dimensions. |
| `clear_range` | `sheet`, `range` | Clear `all`, `contents`, `formats`, `hyperlinks`, or `removeHyperlinks`. |
| `update_sheet` | `action`, `name` | Add, rename, delete, or activate a worksheet. Rename also requires `newName`. |
| `update_workbook` | `calculationMode` | Set `Automatic`, `Manual`, or `AutomaticExceptTables`. |
| `copy_range_to` | `source`, `destination` | Copy `all`, `values`, `formulas`, or `formats`. |
| `read_range_image` | `sheet`, `range` | Return a bounded PNG data URI for visual QA. |
| `read_sheets_metadata` | none | List sheets, IDs, positions, and visibility. |
| `resize_range` | `sheet`, `range` | Set row/column size or request row/column autofit. |
| `update_sheet_view` | `sheet` plus a view field | Set freeze rows/columns, gridlines, zoom, or selected range. |
| `format_range` | `sheet`, `range`, `format` | Apply typed number, fill, font, alignment, dimensions, or borders. |
| `chart` | `action`, `sheet`, `name` | Create, update, or delete a chart. Create also requires `type` and `sourceRange`. |
| `table` | `action`, `sheet`, `name` | Create/delete a table or add/delete table rows. |
| `pivot_table` | `action`, `sheet`, `name` | Create, delete, or refresh a PivotTable. Create needs `source` and `destination`. |
| `save` | none | Explicitly ask Excel to save the current workbook (ExcelApi 1.11). An unsaved workbook uses Excel's native Save As prompt; OfficeKit never supplies a path. |

Every operation is capability-checked inside the Add-in. A missing Office API
returns `unsupported-capability`; it is not replaced with arbitrary script
execution.

## Error handling

| Code | Agent action |
| --- | --- |
| `not-installed` | Run `officekit excel install` only after user approval. |
| `unsupported-platform` | Explain that V1 requires desktop Excel on Windows or macOS. |
| `bridge-unavailable` | Run `officekit excel doctor --json`; inspect its repair guidance. |
| `browser-not-paired` or `session-unavailable` | Open OfficeKit in the intended workbook and connect again, then rediscover sessions. |
| `session-disconnected` or `operation-timeout` | Re-read touched targets before retrying. |
| `unsupported-capability` | Explain the Excel client limitation and wait for a user-selected alternative. |
| `invalid-request` | Repair the request JSON; do not retry unchanged input. |

The request and response limits are intentionally bounded: up to 32 requested
ranges, 50,000 written matrix cells, 20,000 cells across a range read, 25,000
cells scanned by workbook search, 2,500 cells in a screenshot source, 1 MB
request JSON, and 8 MB screenshot payload. Search is literal rather than a
caller-supplied regular expression.
