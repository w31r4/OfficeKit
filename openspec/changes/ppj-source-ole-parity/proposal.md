## Why

OfficeKit already proves one uniquely owned embedded XLSX or DOCX package,
validates replacement bytes, preserves its OLE shell and preview, writes only
the embedded package part, and revalidates the output. PPJ already defines an
`ole.payloadAsset` state field and `setOlePayload` capability, but imported OLE
objects still project as generic opaque elements.

## What Changes

- Materialize a proven embedded XLSX/DOCX package as a content-addressed PPJ
  asset without exposing its OPC locator.
- Project its owner as typed `ole` state with `payloadAsset` and bounded
  `setOlePayload/ole.payload` authority.
- Lower a replacement asset through the existing source-bound OLE writer.
- Keep the shell, preview, relationship identity, position, and all unrelated
  parts source-owned.
- Keep shared, external, ambiguous, unsupported, or oversized OLE graphs opaque.

## Capabilities

### New Capabilities

- `ppj-source-ole-parity`: Capability-issued replacement of one proven embedded
  Office package from imported PPJ state.

### Modified Capabilities

None. The PPJ schema ID and Office wire protocol version remain unchanged.

## Impact

PPJ asset projection, native asset classification, source-bound lowering,
Skill guidance, capability coverage, and one existing OLE contract are
affected. The native OLE readers/writers and PPJ `ole` schema are reused.
