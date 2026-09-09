## 1. Native and PPJ lifecycle

- [x] 1.1 Add optional native bend state, strict literal geometry read/write, PPJ schema/authored/projection/source authority; verify a focused connector test covers explicit values/zero/absence, independent source edit/removal, type switch, preserved bindings/ZIP parts and opaque unsupported graphs.

## 2. Preview and discoverability

- [x] 2.1 Add explicit internal nondefault-bend preview fallback and its regression; update Help, registry, one shapes reference, backlog and coverage, regenerate bindings/docs and pass proto, focused native/JS, reference and strict OpenSpec checks.

Native evidence: SDK 8.0.128 connector/preview filter passes 69/69, zero skipped. The original-source fixture proves explicit zero/negative/removal, raw adj1 values, preserved bindings and SlidePart-only mutation; computed formulas and nonrepresentable rotated bend axes stay opaque with byte-identical no-op.

Closing checks pass: proto:check (lint and regenerated bindings match the reviewed staged output), focused SVG nondefault-bend fallback, preview input/coverage, generated PPJ manual and matrix checks, portability (255), reference sync (333), strict OpenSpec and whitespace. NativeAOT release build and PowerPoint appearance were not performed.
