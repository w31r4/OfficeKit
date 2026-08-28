# Six-sample programmable import evidence

This directory contains compact, reproducible evidence for six real PPTX
inputs.  The inputs themselves are intentionally kept in ignored
`tmp/reference-pptx-downloads/`; see that directory's `SOURCES.md` for
provenance and licensing terms.

Run from a checkout with the reference inputs available:

```sh
node scripts/pptx-six-sample-import.mjs
```

The runner verifies the frozen source hashes, counts visible slide roots and
their recursive group children, requires a byte-identical no-op export, builds a source-bound design profile,
and performs one fresh text edit, placement edit, same-format image replacement,
and source-slide reuse per sample. Where a table has a writable cell, it also
proves one cell-text edit; otherwise that operation is reported as blocked
rather than skipped silently. It re-imports every result and requires the
expected mutation to stay within the target slide XML (plus the replacement
image part and relationship when an image changes). When a rich DrawingML table
or grouped shape is outside the semantic profile, the runner exposes safe
existing text tokens as bounded native leaves and proves one token-splice edit
without rebuilding the object.
`inspect({ kind: "importObject", includeNested: true })` exposes a stable
shape-tree path for every semantic group child.  Children inside a group that
cannot be safely projected are emitted as `opaque-preserved` records derived
from the preserved source subtree; they are discoverable for audit but carry no
edit operation.  `evidence.v1.json` contains hashes, counts, capabilities, and
statuses only; it does not contain source bytes or extracted content. Imported pictures that
carry an Office SVG fallback are surfaced as one image with a separately bound
SVG asset; safe SVG leaves can be edited without replacing the primary raster
relationship.  Unsupported SVG topology remains preserved and reports no
editable leaves.

The current six-sample pass also exercises the newly proven imported-edit paths:
two native fill-color edits, one SVG style edit, and one text edit on a slide with
an existing animation graph. A `blocked` status means that the source sample
does not contain a safe leaf of that kind; it is retained as evidence rather than
treated as a skipped or successful edit.

`render-evidence.v1.json` records the one-pass LibreOffice → Poppler check for
the same six inputs and a bounded placement edit. It includes per-slide PNG
hashes so non-target pages can be checked without storing the rendered images.

This is structural/package evidence.  Windows desktop PowerPoint open/save and
playback remain a separate, not-yet-run host validation.
