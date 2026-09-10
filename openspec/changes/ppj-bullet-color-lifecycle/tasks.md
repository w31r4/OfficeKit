## 1. Direct bullet color choice

- [x] 1.1 Complete schema, authored/source lowering, native projection/validation and choice-local color mutation; verify text/shape marker lifecycle, exact authority, RGB precision/theme/grammar behavior and unknown source preservation with focused native tests.

## 2. Shared behavior and discoverability

- [x] 2.1 Update Help, registry, generated manual/matrix, text guidance and F-03 backlog; verify RGB/alpha preview and theme/follow-text diagnostics, related paragraph/list/table/master regressions, generated checks, portability/reference sync and strict OpenSpec validation.

Validation (2026-09-10): focused tests pass 4/4 after reproducing rejected source color edits and projection failure on invalid native theme tokens. They cover authored/external and imported/embedded picture markers alongside character/number markers, exact alpha and choice presence, independent source lifecycle, XML/ZIP preservation, theme identity, grammar precedence with standalone token-kind validation, and malformed native color/alpha preservation. The shared run passed 33 cases and found one table assertion expecting the previous rounded hex color; it now verifies RGB plus exact alpha 0.7, and its isolated rerun passes 1/1. All 34 selected cases therefore have passing evidence, zero skips. Direct RGB preview retains zero/one/0.12345 alpha; theme/follow-text color is explicitly unavailable. Generated Help/manual/matrix checks, portability/reference sync and strict OpenSpec pass. No full repository test, NativeAOT rebuild or PowerPoint host glyph/layout acceptance was performed.
