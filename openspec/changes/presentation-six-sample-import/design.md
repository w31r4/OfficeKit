# Design

The evidence runner uses only the public OfficeKit import/export API and an
independent ZIP/package reader:

1. Read each ignored reference file with a size bound and verify its frozen
   SHA-256.
2. Count slide roots independently, import the bytes, and compare the public
   `inspect({ kind: "importObject" })` index with that count.
3. Export without edits and require byte identity.
4. Build a source-bound design profile, then run one fresh text edit, one fresh
   placement edit, and one fresh source-slide duplication.  Re-import each
   result and compare package entries; the only expected mutation is the target
   slide XML.
5. Emit compact JSON containing hashes, counts, capabilities, and statuses.

Imported-object inspection remains top-level by default for compatibility.  A
caller that needs a complete visual inventory opts into
`inspect({ kind: "importObject", includeNested: true })`; the result adds
source-tree paths and parent IDs for semantic group children.  When a group is
preserved as an opaque native subtree, its direct visible children are indexed
from the preserved XML as `opaque-preserved` records with synthetic,
revision-bound locators.  These records are audit targets only and never imply
an edit capability.

The runtime remains layered: semantic objects are used when available,
controlled native leaves are used only when the source locator and hash prove
the target, and unknown parts remain opaque.  The runner is diagnostic evidence,
not an edit receipt; package bytes and re-imported values are its authority.

The six inputs live under ignored `tmp/reference-pptx-downloads/`.  Their
provenance and licensing boundary are documented in that directory's
`SOURCES.md`; no source bytes, screenshots, or extracted text are committed.
