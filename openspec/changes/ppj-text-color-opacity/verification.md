# Verification

Delivered bounded PPJ text-color opacity for ordinary text runs and default
text styles, including table default text. The native reader recognizes only a
direct RGB or theme color with at most one direct `a:alpha` child. Projected
third-party formatting remains source-owned without an issued capability.

Verified on 2026-08-30:

- strict PPJ schema and capability-registry JSON parse;
- Release build of `OfficeKit.Codec` with zero warnings and errors;
- the existing integrated PPJ compile/import/project/source-refusal test;
- Open XML validation inside that integrated sample;
- protobuf lint, deterministic generation, and generated-source diff;
- Presentation Skill maintainer consistency check;
- strict OpenSpec validation.

No full suite, package gate, or release candidate was run for this bounded PPJ
2.0 slice.
