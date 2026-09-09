## 1. Upright lifecycle

- [x] 1.1 Infer upright removal from old/new source styles in text/shape/owner-local placeholder and structured table paths, retaining explicit false and existing guards. Verify actual native attribute deletion, source no-op, true/false restoration, fresh projection, unchanged surrounding text/body properties and non-target ZIP parts with focused regressions.

## 2. Public contract

- [x] 2.1 Update schema description, Help, registry, text reference, backlog and coverage, regenerate references/matrix and verify focused native, preview diagnostics, portability/reference-sync and strict OpenSpec checks. Record remaining host/layout and unrelated style-removal limits.

Verification (2026-09-10): upright lifecycle 7/7 passes. Related body/placeholder/table selection 103/104 passes, zero skipped; the opaque Editable assertion in TopLevelTablesAuthorImportEditResizeAndFailClosedOnMergedCells (line 6567) also fails in clean detached baseline bbc1065b under tmp/ppj-upright-baseline. SDK 8.0.128 with repository TMPDIR and single-process build. Preview input/capability, portability 255 files, reference sync 333 files, generated reference/matrix, strict OpenSpec and whitespace checks pass. Initial tests reproduced retained upright on text/shapes and a table postwrite mismatch; source field deletion, compact table semantic normalization and restoration now pass. Table XML comparison ignores redundant namespace declarations only; non-target ZIP bytes remain exact. Existing source capabilities, other style removal guards and unsupported-owner boundaries remain. No wire revision, NativeAOT rebuild or host-layout acceptance.
