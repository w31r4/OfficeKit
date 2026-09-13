## 1. Contract and projection

- [x] 1.1 Add the table-cell field-ID path, `setTextField` capability entry,
      and `tableTextFieldId` native leaf to the PPJ schema/registry; verify
      both JSON contracts parse.
- [x] 1.2 Project static table-cell field IDs with the existing cell/field
      ordinals; verify the native reference advertises the exact path.

## 2. Source-bound codec

- [x] 2.1 Accept an ID-only change while preserving field type/cached text and
      reject combined identity/text or automatic changes; verify native
      mutations remain source-bound.
- [x] 2.2 Re-prove the table cell and field ordinal and splice only the direct
      `a:fld/@id` token; verify stale, invalid, and unsupported graphs fail
      closed.

## 3. Evidence and publication

- [x] 3.1 Add an authored/imported table-field identity regression covering
      type/cache/topology preservation, SlidePart-only scope, Open XML
      validity, re-projection, and rejection cases; verify the focused test
      passes.
- [x] 3.2 Update the F-03 backlog, coverage, and presentation Skill reference;
      run OpenSpec strict validation and the narrow codec/portability gates.
