## Why

G-12 in `docs/ppj-svg-preview-gap-audit.zh-CN.md` identifies that local preview can overwrite existing outputs, list PNGs that do not exist, and leave unlabelled partial results. Subsequent rendering fixes need evidence that reliably identifies the input, assets, generated files and failures.

## What Changes

- Require a fresh output directory and exclusive artifact writes; preserve existing directories, files, input PPJ, source PPTX and assets.
- Track output lifecycle explicitly, retaining labelled incomplete results after raster or filesystem failure and recording only successfully written artifacts.
- Record actual SVG/PNG hashes and byte counts, source/compiled input identity, asset hashes and renderer/dependency identity without claiming visual correctness.
- Reuse loaded workspace asset bytes for compilation and SVG embedding to avoid evidence referring to different resource reads.
- **BREAKING**: reusing an existing preview output directory fails instead of overwriting it; unavailable PNG generation cannot return an unqualified successful output result. Existing successful safe page filenames and legacy manifest fields remain compatible where truthful.
- Add deterministic negative tests and a narrow regular test gate, and update the G-12 audit status only after verification.

## Capabilities

### New Capabilities

- `ppj-preview-output-evidence`: Non-destructive preview artifact publication and truthful success/failure evidence.

### Modified Capabilities

None.

## Impact

Targets: `src/ppj/svg-preview.mjs`, a small leaf output helper if needed, `test/ppj-svg-preview.mjs`, a new focused output-evidence test, package/test-gate entries, preview usage documentation and the gap audit. Reuse `workspace.mjs` hash/exclusive-file helpers and the existing optional sharp dependency; no protocol version change, external engine, new dependency or network download.

This is the first bounded change within the full G-01 through G-16 goal. It does not claim to repair element/chart semantics, all G-11 support declarations, G-13 page selection, or the remaining gap list. Honest output evidence is necessary but not sufficient for rendering completion.
