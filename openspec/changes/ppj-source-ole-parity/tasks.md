## 1. Asset and projection

- [x] 1.1 Materialize proven embedded XLSX/DOCX parts as local PPJ assets.
- [x] 1.2 Project eligible owners as typed OLE with bounded payload authority.

## 2. Source-bound lowering

- [x] 2.1 Map replacement PPJ assets to private OLE native asset identities.
- [x] 2.2 Lower only `payloadAsset` and retain the existing native writer proof.
- [x] 2.3 Keep unsupported OLE graphs opaque and fail closed on stale authority.

## 3. Agent surface and lean verification

- [x] 3.1 Regenerate `ppj.md` and update imported-edit guidance and coverage.
- [x] 3.2 Extend one existing OLE test with PPJ projection/edit/reimport.
- [x] 3.3 Run the focused test, Skill maintainer, and strict OpenSpec check once.
- [x] 3.4 Commit atomically and fast-forward main without force pushing.

## Evidence

- `OleWorkbookPayloadReplacementIsValidatedAndGraphBound` passed after the
  existing XLSX OLE contract was extended through PPJ projection, local asset
  materialization, `payloadAsset` replacement, compilation, and second
  projection.
- The PPJ build reported only `ppt/embeddings/native-workbook.xlsx` as changed.
  Every other OPC part, including the OLE shell, slide relationship, preview,
  SmartArt and content-part evidence in the same fixture, stayed byte-identical.
- Second projection recovered the replacement workbook SHA-256 and a fresh
  typed OLE payload binding. The same existing contract continues to reject
  shared packages, changed locators, missing assets and malformed workbooks.
- `presentation-skill-maintainer check` passed with 151 Help APIs, 73 native
  leaves and 13 host-only operations after regenerating `ppj.md`.
- `openspec validate ppj-source-ole-parity --strict` passed.
- No schema or wire change, new writer, new test file, sample matrix, full
  suite, package gate, playback claim, preview regeneration or raw OOXML
  surface was added.
