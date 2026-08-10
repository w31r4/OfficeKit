# Help

`presentation.help(...)` queries the live OfficeKit Help catalog. Treat its
bounded NDJSON as authoritative; this file deliberately does not duplicate
catalog records that would drift from the shipped runtime.

```ts
const result = presentation.help("*", {
  search: "accessibility|shape|chart",
  include: ["index", "examples", "notes"],
  maxChars: 8000,
});
console.log(result.ndjson);
```

- `query` accepts an exact dotted name, a glob such as `shape*`, or `*`.
- `search` is a case-insensitive regular expression over selected records.
- `include` accepts `index`, `examples`, and `notes` as an array or CSV string.
- `maxChars` bounds output; `result.truncated` and the final notice record tell
  the Agent to narrow the query or raise the budget.

Use Help to identify a typed primitive, then inspect/resolve the target, apply
the edit, render, and verify. Full option/result shapes remain in
[`presentation.spec.md`](./presentation.spec.md); domain contracts remain in
the adjacent reference pages.
