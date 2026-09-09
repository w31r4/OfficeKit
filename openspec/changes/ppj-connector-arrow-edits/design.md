## Context

See proposal.md. ApplyLineElement already treats omission as no arrow; connector editing currently only allows frame/stroke/endpoints/accessibility. PptxLineStyleCodec preserves arrow sizes for a retained arrow but requires size state to be empty when its arrow is removed.

## Goals / Non-Goals

Goals: complete the existing connector arrow fields under an explicit source capability.
Non-goals: new arrow shapes, PPJ arrow-size fields, routing changes or wider opaque authority.

## Decisions

- Add setConnectorArrows with startArrow/endArrow fields to schema, validator and editable-connector projection. Do not silently broaden setStroke's advertised fields.
- Change only the requested end. Use existing arrow value normalization; omission or none removes that end and clears its own width/length, while changing a retained arrow keeps those attributes and leaves the opposite end untouched.
- Use the existing semantic writer so source/native frame anchors and site bindings retain their authority. Reject requests lacking the issued capability; do not infer editability from a leaf alone.
- One focused source fixture carries both a frame anchor and nondefault native arrow sizes. Independently add/change/delete arrow fields from its fresh source projection, compare native XML and fresh PPJ, and check the target SlidePart is the only changed ZIP member.

## Risks / Trade-offs

- Removing an arrow while retaining orphan dimensions causes invalid native state → clear only the removed end's dimensions.
- Replacing both ends from defaults could alter untouched source values → apply only changed fields.
- Existing source-owned connector topology may not be modeled → only source.Editable connectors receive the semantic capability.

## Migration Plan

Additive PPJ capability; no wire or source-free behavior change. Publish the isolated change after native and generated documentation checks; do not claim host or complete F-04 validation.

Implementation finding: PPJ open and native arrow were not mapped by the shared PPJ arrow helpers, causing authored validation or lost fresh projection. Map open to arrow on authored/source lowering and arrow to open on projection; retain the existing native vocabulary. This is required for the declared arrow enum, not a new arrow shape. The same helper is shared by existing PPJ line/table paths.
