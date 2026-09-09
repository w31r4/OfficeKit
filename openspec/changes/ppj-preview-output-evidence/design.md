## Context

The existing `src/ppj/svg-preview.mjs` writes page SVGs with `writeFile`, PNGs through sharp `toFile`, and a manifest after all pages. It reuses output directories and reads asset paths again after `loadPpjWorkspace` has already loaded them. Missing sharp adds a diagnostic after the old overall status is calculated, but every page still gets a PNG filename in the manifest.

`src/ppj/workspace.mjs` already supplies `sha256` and `writeExclusiveFile` (temporary file plus exclusive hard link). `src/ppj/render-review.mjs` reserves a fresh directory for external rendering, but recursively removes its output on error; local preview needs retained failure evidence for repair instead. `src/renderers/sharp.mjs` demonstrates buffer-based raster output and optional lazy loading.

The existing preview smoke passes an already-created mkdtemp directory as output. It must instead create a temporary parent and select an unused child destination. That intentional compatibility change follows the repository's input-preservation rule.

## Goals / Non-Goals

**Goals:** Close G-12 through exclusive writes, truthful artifact accounting, explicit failure lifecycle and a single input/asset snapshot. The implementation stays a local leaf module and uses small deterministic fault-injection tests.

**Non-Goals:** No element/chart drawing rewrite, generalized scene graph, new external renderer, dependency download, page selection, visual scoring or global capability-declaration repair. Those remain in the full G-01–G-16 goal; G-05 asset snapshot and G-11 dependency-status bugs receive only the changes necessary for correct output evidence.

## Decisions

### 1. Fresh directory with exclusive files, not overwrite or cleanup

Resolve outputDir relative to the invocation cwd. Create its parent as needed, reserve the destination with non-recursive mkdir, and reject EEXIST for directories, files and symlinks. Compare canonical protected input/source/asset locations where applicable. Use `writeExclusiveFile` for all published artifacts, including PNG buffers; never let rasterization write arbitrary filesystem destinations.

Derive safe, unique leaf filenames before page publication. Preserve existing safe page stems such as `page-claim`; map unsafe/reserved/colliding IDs to deterministic ordinal-based names, checking uniqueness across case-folded stems. Preserve original IDs separately in evidence. Do not use arbitrary PPJ IDs directly as paths.

Alternative: reuse an empty directory. Rejected because proving exclusive ownership of a pre-existing directory is weaker and can mix evidence from different runs. No `--force` option is added.

### 2. Retain explicitly incomplete runs

After destination reservation, write `render.pending.json` exclusively with input identity and an incomplete state before page files. If that initial write fails, fail without attempting pages. Attempt SVG/PNG publication with stage- and page-specific error capture; keep successfully published files.

Publish `render.json` exclusively at the end with `output.status: complete|incomplete`, actual artifacts and failed attempts. Complete means requested output files were produced, not that drawing is correct. Remove only this run's pending marker after the final manifest is successfully published by the existing helper. An interrupted run or a final-manifest failure retains pending evidence and cannot have a valid complete manifest.

A final manifest is authoritative when present; a leftover pending marker after final publication is harmless but must not be interpreted as overriding the final state. On failure to remove that marker, report a cleanup diagnostic without deleting output artifacts. This is publication atomicity, not an fsync/power-loss durability guarantee.

For incomplete publication, throw a structured error after writing obtainable evidence, attaching output path and receipt for programmatic callers. The existing CLI's error path supplies nonzero exit; no exception should be converted to a successful return. If manifest writing itself fails, surface both the original stage failure and evidence-write failure as available. Do not recursively delete the destination.

Alternative: stage everything and rename a directory. Rejected for this small change because cross-platform rename-to-existing behavior and total-output buffering complicate preservation and discard useful partial repair evidence.

### 3. Truthful versioned receipt with compatible successful fields

Add a manifest schema such as `office-kit/ppj-svg-preview-output/v1`. Preserve the renderer name, canvas, diagnostics, existing compile hashes and successful page-level `file`/`png` references. Add:

- input raw hash/path, canonical program hash, compiled candidate hash, sourceBound and actual source hash when present;
- asset id/hash/byte-size records from the loaded byte snapshot;
- output directory, production state, failed attempts and `artifacts[]` with pageId/kind/file/sha256/bytes;
- preview implementation identity (version plus source digest), Node version and actual sharp/libvips version information when loaded;
- an explicit statement that the evidence is a local SVG preview and visual correctness still needs review.

Hashes and byte sizes are computed from the exact buffers published, not reread original paths. Keep output execution state separate from the existing element-support status so G-12 does not claim that inaccurate drawing declarations are fixed. Dependency-unavailable output must nevertheless never retain a supported/successful publication verdict.

Alternative: only omit the missing PNG field. Rejected because it does not record why an output is incomplete or bind artifacts to inputs.

### 4. One loaded resource snapshot

Build the SVG asset map from `workspace.assets[].data`, the same bytes passed by `compilePpjWorkspace`, and record their actual hashes. Preserve initial resource validation and normal codec rejection. Do not reopen asset URIs after compilation. This removes the second-read window without changing image-fit, masking or effects behavior.

Missing or inconsistent asset IDs at the rendering boundary produce explicit unavailable evidence; they do not become an empty href with a successful image claim. Source and raw PPJ bytes are also taken from the existing workspace snapshot.

### 5. Lazy raster backend with deterministic test seams

Load sharp only when publishing requested PNGs; obtain PNG bytes before exclusive publication. Separate the small publication helper from page drawing, with narrowly scoped injected raster/write operations for tests. If module loading fails, preserve SVG output, emit unavailable dependency evidence and fail requested PNG publication. Raster execution failure is a different failed stage, not automatically a missing dependency.

Do not add a root import or change protocol bindings. The existing no-outputDir in-memory API stays SVG-only and dependency-lazy. Use a real-sharp smoke in addition to injected tests so test stubs cannot be the only successful path.

## Risks / Trade-offs

- [Existing users reuse directories] → Fail with a clear destination-exists message; document choosing a fresh child path and update the current smoke.
- [Failure evidence cannot be written due to filesystem failure] → Pending evidence is written first; never leave a complete manifest, surface the evidence-write error and retain obtainable artifacts. No promise can override a wholly unwritable filesystem.
- [Partial output consumes disk] → Document retained output paths; only users or a separately authorized cleanup remove them. Avoid deleting inputs or broad directories.
- [Symlink or case-insensitive filenames] → Refuse existing destination symlinks, use a checked leaf namespace and test collision/path cases; do not claim protection from an adversarial process replacing filesystem ancestors during the run.
- [Manifest fields drift from visual diagnostics] → Separate production state from content support; G-11 remains explicitly open.
- [Concurrent edits to unrelated codec work] → Limit modifications to preview/output tests and its documentation/gate entries; no logBase or native code changes.

## Migration Plan

First add fault-injection and preservation tests, then update publication and resource snapshot use. Move the existing smoke to a fresh child directory; keep existing safe page filenames and successful compatibility fields. Register the small output-only test in the regular gate and run the actual codec/sharp smoke separately. Update the gap audit with concrete commands and remaining boundaries. No automatic route switch, release, protocol migration or commit/push is part of this change.
