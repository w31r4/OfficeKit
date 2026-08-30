## ADDED Requirements

### Requirement: Proven OLE payload projects as typed PPJ state
An imported OLE object SHALL project as typed `ole` state only when the codec
has proved one supported uniquely owned embedded Office package.

#### Scenario: Agent discovers an embedded workbook or document
- **WHEN** the source owner has an issued XLSX or DOCX payload binding
- **THEN** PPJ contains a content-addressed local asset, `payloadAsset`, and
  bounded `setOlePayload/ole.payload` authority

#### Scenario: OLE graph is outside the safe profile
- **WHEN** the payload is shared, external, ambiguous, unsupported, or invalid
- **THEN** the element remains opaque and no payload capability is issued

### Requirement: Declarative payload replacement
Changing an eligible PPJ OLE element's `payloadAsset` SHALL request replacement
of only the bound embedded package.

#### Scenario: Agent selects a valid same-kind asset
- **WHEN** the new content-addressed asset has the required XLSX or DOCX MIME
  and the nativeRef remains unchanged
- **THEN** build writes only the embedded package part and second projection
  recovers the replacement asset digest

#### Scenario: Agent changes package kind or OLE shell state
- **WHEN** the request changes content type, preview, display mode, ProgID, or
  any source-owned field
- **THEN** build rejects before writing output

### Requirement: OLE shell and presentation remain source-owned
Payload replacement SHALL retain the OLE frame, preview, relationship, slide,
and every unrelated package part.

#### Scenario: One payload changes
- **WHEN** a valid replacement succeeds
- **THEN** only the embedded package part differs and review reports that the
  source preview image was not regenerated
