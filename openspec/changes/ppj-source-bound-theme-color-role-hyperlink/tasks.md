## 1. Contract

- [x] 1.1 Add the hyperlink capability and field to the PPJ schema, registry, coverage entry, and generated presentation reference; verify JSON parsing and Skill reference synchronization.
- [x] 1.2 Validate the OpenSpec artifacts and confirm no protobuf field or protocol version changes are needed.

## 2. Codec

- [x] 2.1 Project the strict six-role RGB observation with `setThemeColorRoleHyperlink` while retaining existing color-role authority; verify unsupported topology stays source-owned.
- [x] 2.2 Extend source-bound compilation and the PPTX writer to patch only `a:hlink/a:srgbClr/@val`; verify the existing one-theme-field mutation boundary remains enforced.

## 3. Regression

- [x] 3.1 Add a focused native regression covering no-op bytes, changed-part scope, non-target ZIP preservation, Open XML validity, second projection, and rejection of deletion or tampered authority.
- [x] 3.2 Run focused hyperlink tests, theme regressions, OpenSpec strict validation, schema/registry checks, protobuf lint/generation, and Presentation Skill maintenance; record any environment-only skip (`node test/reference-skills.mjs` cannot start `pdftoppm` in this environment).
