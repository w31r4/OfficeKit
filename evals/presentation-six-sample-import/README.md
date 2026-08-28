# Six-sample programmable import evidence

This directory contains compact, reproducible evidence for six real PPTX
inputs.  The inputs themselves are intentionally kept in ignored
`tmp/reference-pptx-downloads/`; see that directory's `SOURCES.md` for
provenance and licensing terms.

Run from a checkout with the reference inputs available:

```sh
node scripts/pptx-six-sample-import.mjs
```

The runner verifies the frozen source hashes, counts visible slide roots,
requires a byte-identical no-op export, builds a source-bound design profile,
and performs one fresh text edit, placement edit, and source-slide reuse per
sample.  It re-imports every result and requires the expected mutation to stay
within the target slide XML.  `evidence.v1.json` contains hashes, counts,
capabilities, and statuses only; it does not contain source bytes or extracted
content.

`render-evidence.v1.json` records the one-pass LibreOffice → Poppler check for
the same six inputs and a bounded placement edit. It includes per-slide PNG
hashes so non-target pages can be checked without storing the rendered images.

This is structural/package evidence.  Windows desktop PowerPoint open/save and
playback remain a separate, not-yet-run host validation.
