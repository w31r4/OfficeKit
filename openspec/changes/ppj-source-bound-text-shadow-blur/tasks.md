## 1. PPJ contract

- [x] 1.1 Add `textShadowBlurRadiusEmu` to the native leaf schema and capability registry, pointing to `run.style.shadow.blur`, and verify JSON schema/registry parsing succeeds
- [x] 1.2 Update presentation reference, coverage, and gap backlog entries for the bounded direct-run shadow blur profile, and verify the capability matrix is regenerated and current

## 2. Native projection

- [x] 2.1 Project a strict direct rich-text `a:outerShdw/@blurRad` as `textShadowBlurRadiusEmu` while retaining the existing semantic `shadow.blur`, and verify valid authored/imported runs expose the exact EMU value
- [x] 2.2 Keep missing, transformed, multi-effect, or malformed outer-shadow graphs opaque, and verify the focused regression asserts no shadow blur leaf for each unsupported topology

## 3. Source-bound edit

- [x] 3.1 Add edit-plan validation, proof, read, and dispatch support for `textShadowBlurRadiusEmu` with a changed bounded EMU token, and verify invalid/no-op values fail closed
- [x] 3.2 Token-splice only the owning run `outerShdw/@blurRad` attribute while preserving other XML and package parts, and verify changed-parts and Open XML assertions pass

## 4. Regression and gates

- [x] 4.1 Add an authored → source-bound edit → second projection regression covering semantic/native values, exact token, non-target ZIP preservation, and unsupported topology
- [x] 4.2 Run the focused NativeAOT test, OpenSpec strict validation, presentation matrix/Skill sync checks, and portability/reference gates; record any environment-only skip
