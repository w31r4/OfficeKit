## 1. OpenSpec and wire surface

- [x] 1.1 Document the bounded static field-type lifecycle and evidence limit.
- [x] 1.2 Add the capability operation, field path, and native leaf kind to
      the PPJ schema and semantic validator.

## 2. Codec implementation

- [x] 2.1 Issue `textFieldType` leaves for static fields and advertise
      `setTextField` on eligible text owners.
- [x] 2.2 Allow only static field-type changes in source-bound text compilation
      and keep field IDs, display text, automatic state, and topology fixed.
- [x] 2.3 Re-prove and token-splice only `a:fld/@type` in the owning slide.

## 3. Evidence and publication

- [x] 3.1 Add focused authored/imported lifecycle regression coverage.
- [x] 3.2 Update PPJ schema description, capability registry, presentation Skill,
      preview diagnostics, and the gap backlog.
- [x] 3.3 Run focused codec/preview/protocol/portability gates, validate this
      change strictly, then make one atomic commit and push `main`.
