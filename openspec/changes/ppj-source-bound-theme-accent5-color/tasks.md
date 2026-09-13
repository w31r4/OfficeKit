## 1. Contract

- [x] 1.1 Add the accent5 capability and field to the PPJ schema, registry, coverage entry, and generated presentation reference; verify JSON parsing and Skill reference synchronization.
- [x] 1.2 Validate the OpenSpec artifacts and confirm no protobuf field or protocol version changes are needed.

## 2. Codec

- [x] 2.1 Extend the theme projection and semantic capability map with `setThemeAccent5Color` for the observed `accentColors.accent5` field; verify strict source-bound ownership rules remain fail-closed.
- [x] 2.2 Extend source-bound compilation and the PPTX writer to patch only `a:accent5Color/a:srgbClr/@val` while retaining all other accent slots and theme bytes.

## 3. Regression

- [x] 3.1 Add a focused native regression covering no-op bytes, accent5 changed-part scope, non-target ZIP preservation, Open XML validity, second projection, and rejection of unowned edits or tampered authority.
- [x] 3.2 Run the focused native tests, theme regressions, OpenSpec strict validation, schema/registry checks, protobuf lint/generation, and presentation Skill maintenance check.
