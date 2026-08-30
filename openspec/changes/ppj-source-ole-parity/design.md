## Context

An imported OLE object receives `ole_workbook` only for one uniquely bound
ordinary XLSX package, or `ole_office_package` only for the current bounded
DOCX profile. Those bindings retain the private part path, relationship ID,
content type, source digest, and package kind. The original payload bytes stay
in the preserved source PPTX snapshot rather than `ArtifactEnvelope.assets`.

PPJ's public asset model is content-addressed and local. The projector must
therefore materialize the proven package into the PPJ asset directory while
also supplying the private native-purpose asset identity expected by the
existing writer during source compilation.

## Goals / Non-Goals

**Goals:**

- Let an Agent inspect, retain, or replace a proven embedded XLSX/DOCX through
  ordinary PPJ asset state.
- Keep source package locators and OLE XML out of PPJ.
- Reuse exact content-type, hash, relationship, package-validation, and
  postwrite checks already owned by the codec.
- Preserve every unrelated package byte.

**Non-Goals:**

- Authoring a new OLE object or changing display mode, ProgID, preview, shell,
  relationship, package kind, or content type.
- Editing the embedded workbook/document semantically inside PPJ.
- Supporting legacy binary Office formats, external links, shared payloads, or
  arbitrary embedded file containers.
- Inferring visual updates to the OLE preview image.

## Decisions

### 1. The payload is an ordinary PPJ asset

Projection extracts only the proven embedded package entry from the immutable
source PPTX, verifies it against the opaque-part digest, and emits a local
content-addressed `.xlsx` or `.docx` asset. The typed OLE element points to that
asset. The PPJ never carries package bytes inline.

### 2. Native-purpose identity remains private

The projector also makes the same bytes available to the source compiler under
the private `asset/presentation/ole-*` identity required by `PptxAssetCatalog`.
The public program sees only its stable PPJ asset ID and relative URI.

### 3. Replacement is complete payload state, not an operation log

Changing `payloadAsset` requests one complete replacement. The compiler
requires the exact element nativeRef and `setOlePayload/ole.payload`, maps the
new PPJ asset to its private native-purpose ID, and sets only the existing
replacement field on the fresh wire binding. The native writer independently
validates the package and source relationship before writing.

### 4. The source asset remains immutable

The original projected asset declaration cannot be removed or mutated during a
source-bound compile. A replacement is a new declared asset. Builds never
overwrite either the source PPTX or the extracted payload file.

## Risks / Trade-offs

- [Large context from embedded packages] -> PPJ stores only metadata and a
  relative URI; bytes stay in the asset directory.
- [Preview becomes stale] -> The preview remains source-owned and review must
  disclose that payload replacement does not regenerate it.
- [Wrong package type] -> MIME, content-addressed native ID, Open XML package
  profile, and exact source binding all fail closed.
- [Opaque package extraction expands authority] -> Only bindings already issued
  by the native OLE catalog can materialize or advertise replacement.

## Migration Plan

No migration. Fresh projections of eligible OLE objects become typed. Existing
PPJ and unsupported imported objects remain valid and opaque.

## Open Questions

None.
