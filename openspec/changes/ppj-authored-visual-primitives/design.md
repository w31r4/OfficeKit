## Context

The current PPJ schema exposes 534 documented fields and 14 element types. PPTD
has fewer element kinds but a much longer chart/style specification. The visible
difference therefore comes from compiler depth and documentation, not from JS
being Turing-complete or PPJ being JSON.

The current authored compiler already lowers text, preset shapes, solid fills,
images/background images, basic charts, tables, connectors, groups, motion, and
deck metadata. The native codec already owns bounded custom geometry. Missing
visual states fall into two classes:

1. schema fields with an existing native owner but no PPJ lowering;
2. schema fields that need a new additive native semantic owner.

## Decisions

### Reject silent visual loss

Every authored style block is consumed property-by-property. Unsupported present
fields throw `unsupported_ppj_compile_feature` before a PPTX is returned. Omitted
fields keep their existing defaults. This turns the capability registry into a
truthful availability index rather than a list of aspirations.

### Reuse the existing custom-geometry codec

PPJ custom coordinates are normalized from `viewBox` into an integer DrawingML
path viewport. The lowerer emits move, line, quadratic, cubic, and close commands
into `PresentationCustomGeometryPath`; the existing native codec remains the
single validator/writer. Preset adjustment arrays remain fail-closed until they
have named-guide semantics.

### Add typed paint, not raw XML

Gradient and line-alpha support use additive protobuf messages with closed stop,
angle, and opacity fields. PPJ never accepts DrawingML tokens, transforms, paths,
or relationship IDs. Existing imported gradients remain source-owned unless a
separate capability is proven.

### Expand charts and tables by semantic value

First own existing PPJ fields: legend placement, stacking, gap width, axis and
gridline visibility, chart/plot fills, cell fill/text/borders. Richer PPTD-only
chart families and per-point effects require separate typed designs and are not
smuggled into generic property bags.

## Verification

Use one existing PPJ integration test containing custom geometry, gradient/alpha,
chart style, and table style. Add only regressions for real failures. Run the
narrow codec test, generated-reference check, proto check, and strict OpenSpec
validation; defer the full suite until the larger PPJ 2.0 release gate.
