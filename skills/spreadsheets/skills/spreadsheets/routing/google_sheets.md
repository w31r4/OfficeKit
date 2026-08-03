## Google Sheets-targeted output

Create and verify a local `.xlsx` with the Spreadsheets Skill first. OfficeKit
does not upload files, create cloud spreadsheets, or operate a Google Drive.
After the local workbook passes semantic and render QA, the user or another
host may import it into Google Sheets.

For an existing native Google Sheet, obtain the requested data or export from
the host and operate on the resulting local file. Do not silently replace a
live cloud edit with an unrelated workbook mutation.

Return the verified `.xlsx` path, SHA-256, and evidence envelope. If the user
needs a cloud link, state that the import is a separate host step.
