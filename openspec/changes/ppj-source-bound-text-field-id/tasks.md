## 1. PPJ contract

- [x] 1.1 Add `textFieldId` to projection, strict value normalization, schema,
  registry, generated matrix, and public presentation references; verify the
  registry and schema parse and the matrix check passes.
- [x] 1.2 Add source-bound direct-field proof, readback, and one-token
  `a:fld/@id` patching; verify stale, invalid, automatic, and unchanged cases
  fail closed or remain no-ops.

## 2. Regression

- [x] 2.1 Add an authored-to-imported lifecycle test that removes embedded PPJ,
  edits one field ID, recompiles, and projects again; verify type, cached text,
  topology, and non-slide bytes are preserved.

## 3. Evidence and publication

- [x] 3.1 Update the gap backlog, coverage entry, and presentation Skill with
  the field boundary and evidence limit; verify generated/reference sync.
- [x] 3.2 Run strict OpenSpec validation, focused codec tests/build, and the
  narrow PPJ schema/matrix/portability gates before staging one atomic commit.
