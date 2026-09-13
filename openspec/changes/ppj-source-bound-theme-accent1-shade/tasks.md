## 1. Contract

- [x] 1.1 Add the accent1 shade capability and field to the PPJ schema, registry, coverage entry, backlog, and generated presentation reference; verify JSON parsing and Skill reference synchronization.
- [x] 1.2 Validate the OpenSpec artifacts and confirm no protobuf field or protocol version changes are needed.

## 2. Codec

- [x] 2.1 Import the strict direct accent1 RGB plus one shade observation and project `setThemeAccent1Shade`; verify unsupported transform topology stays source-owned.
- [x] 2.2 Extend source-bound compilation and the PPTX writer to patch only `a:shade/@val`; verify one-theme-field and capability boundaries remain enforced.

## 3. Regression

- [x] 3.1 Add a focused native regression covering projection, no-op bytes, changed-part scope, non-target ZIP preservation, Open XML validity, second projection, and rejection of deletion, combined edits, unsupported siblings, or tampered authority.
- [x] 3.2 Run focused shade tests, theme regressions, OpenSpec strict validation, schema/registry checks, protobuf lint/generation, and Presentation Skill maintenance; record the environment-only `node test/reference-skills.mjs` skip if `pdftoppm` is unavailable.

> Validation note: `node test/reference-skills.mjs` remains environment-only blocked because `pdftoppm` is not installed; the focused shade/tint tests, schema/registry, protobuf, OpenSpec, coverage, and Presentation Skill checks pass.
