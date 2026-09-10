## 1. Contract and projection

- [x] 1.1 Add the table-cell field path, `setTextField` capability entry, and
      `tableTextFieldType` native leaf to the PPJ schema/registry and verify
      the JSON contract parses.
- [x] 1.2 Project static field-type leaves from eligible fixed-topology table
      cells and advertise the exact capability; verify a source projection
      contains the cell and field binding.

## 2. Source-bound codec

- [x] 2.1 Accept a type-only change without routing it through cached-text
      replacement, while rejecting combined text/type or identity changes;
      verify native mutations remain source-bound.
- [x] 2.2 Re-prove the table cell and field ordinal and splice only its direct
      `a:fld/@type` token; verify automatic, stale, invalid, and unsupported
      table graphs fail closed.

## 3. Evidence and publication

- [x] 3.1 Add an authored/imported table-field lifecycle regression covering
      ID/cache/topology preservation, SlidePart-only scope, Open XML validity,
      re-projection, and rejection cases; verify the focused test passes.
- [x] 3.2 Update the F-06 backlog, coverage, presentation Skill reference,
      and change task state; verify OpenSpec strict validation, proto check,
      maintainer/portability checks, and the focused codec regression pass.
