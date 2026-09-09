## 1. Anchor-center presence lifecycle

- [x] 1.1 Add and validate the native deletion command, apply/normalize it, and infer PPJ removal; verify shared lifecycle tests across supported owners, compact restoration, native command rejection and focused body/AutoFit/placeholder/TableText tests.
- [x] 1.2 Regenerate protocol bindings and public field descriptions; verify proto:check, preview input/capability, generated reference/matrix, Skill portability/reference sync, strict OpenSpec and whitespace checks; record codec compatibility and host-layout limits in coverage.

Verification (2026-09-10): native lifecycle/body/AutoFit/placeholder/TableText 112/112, zero skipped, SDK 8.0.128. C# and JS preserve field-32 boolean bytes; field-40 marker round-trip passes. proto:check passes against regenerated/staged binding. Scene wire, preview input/capability, generated reference/matrix, maintenance, portability (255 files), reference sync (333 files), strict OpenSpec and whitespace pass. Updated codec required; no NativeAOT rebuild or host layout acceptance. Full backlog remains open.
