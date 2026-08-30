## 1. Native Picture Contract

- [x] 1.1 Add additive custom-mask path state to `PresentationImage` and regenerate bindings.
- [x] 1.2 Reuse the native custom-geometry validator/writer for authored picture masks.
- [x] 1.3 Read only canonical literal custom picture geometry and keep irregular masks opaque.

## 2. PPJ Compiler and Projection

- [x] 2.1 Lower PPJ custom image masks into the shared wire path graph.
- [x] 2.2 Project canonical custom image masks back into PPJ and reject source-bound topology mutation.

## 3. Evidence and Agent Surface

- [x] 3.1 Extend the existing comprehensive PPJ round-trip contract with custom-mask native and fail-closed evidence.
- [x] 3.2 Update the capability registry, media/layers guidance, generated PPJ manual, coverage, and OpenSpec state.
- [x] 3.3 Run only the owning C# contract, proto consistency, Skill registry, and strict OpenSpec checks.
