## 1. Contract and lowering

- [x] 1.1 Add `textStyle.softEdge` to the PPJ schema and add run/default-run
  wire fields.
- [x] 1.2 Parse soft edge on authored run and default-run styles.
- [x] 1.3 Lower soft edge after the other bounded text effects while preserving
  the direct effect owner boundary.

## 2. Evidence and bookkeeping

- [x] 2.1 Add a minimal authored rich-text XML and embedded-recovery test.
- [x] 2.2 Update F-03 coverage, capability notes, and generated presentation
  Skill reference.
- [x] 2.3 Run strict OpenSpec validation, proto/schema/reference checks, the
  focused native test, and diff checks.
