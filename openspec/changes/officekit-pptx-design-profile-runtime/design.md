## Design

`Presentation.designProfile()` calls the existing bounded semantic inspection
and derives stable counts and signatures from the model. Imported state supplies
only the source package SHA-256; no source bytes or physical paths enter the
result. Component candidate IDs are included only as defensive summaries when
requested, and all mutation operations still require their own current
source-bound preflight.

The profile is deliberately descriptive. A layout family, repeated component,
or color occurrence can guide an Agent's plan, but it does not imply that the
object can be cloned, edited, or authored. Existing `reuseSourceSlide`,
`reuseSourceComponent`, and native-leaf APIs remain the only mutation paths.

All arrays are deterministically ordered and bounded by `maxItems` (1–256).
Unsupported or unavailable candidate inspection is returned as an explicit
reason rather than an empty success that could be mistaken for “no candidates”.

## Verification

- Source-free profile: no revision binding, deterministic canvas/type/density
  evidence, and no candidate authority.
- Imported profile: exact source SHA-256, opaque-object summary, candidates,
  and byte-identical no-op export.
- `includeComponentCandidates: false` reports that candidates were disabled.
- No output contains source paths, XML selectors, raw XML, or source bytes.
