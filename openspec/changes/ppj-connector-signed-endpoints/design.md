## Context

See proposal.md. The wire uses signed int64 endpoint fields and DrawingML xfrm offsets are signed. The current resolver, native endpoint validation and import return condition impose nonnegativity. Native catalog placement proof separately rejects negative connector offsets. Ppj frame coordinates already accept signed numbers.

## Goals / Non-Goals

Goals: preserve meaningful signed connector coordinates across the same compiler/import/edit path and bound arithmetic explicitly.
Non-goals: arbitrary negative placement for other element types, native site evaluation, routing, host acceptance or production renderer switching.

## Decisions

- Use DrawingML ST_Coordinate bounds -27273042329600..27273042316900 EMU and ST_PositiveCoordinate extents 0..27273042316900. Validate endpoints and their differences before building xfrm. Do not narrow all coordinates to Int32 merely to imply a host guarantee: the library contract remains schema-level. Microsoft documents narrower historic PowerPoint restrictions separately: https://learn.microsoft.com/en-us/openspecs/office_standards/ms-oe376/65893f0f-b482-488f-afd2-af4023d1f9b0 . Standard bounds are also specified at https://learn.microsoft.com/en-us/openspecs/office_standards/ms-oe376/4f890b34-61b8-4d22-beb7-77ac953e66a8 . No host-range compatibility claim follows from this change.
- Share coordinate/frame predicates between native writer, reader and direct connector placement proof. Bound native offset/extents before adding them and recheck rotated endpoints. Unknown or out-of-range native topology remains opaque.
- Retain existing from/to syntax and EMU rounding. Negative literal or resolved object coordinates are not clamped; PPJ schema documentation explains the signed parent space.
- Reuse existing connector tests: make the rotated cross-group fixture positive, then verify signed literal/source endpoint edits and unchanged package members with a fresh projection.

## Risks / Trade-offs

- Imported large offsets can overflow intermediate arithmetic → reject them before adding extents or rotating.
- A broader catalog proof can accidentally enable unrelated topology → restrict signed placement changes to connector kind and keep existing identity/structure checks.
- Rotation may place an endpoint outside the supported range → preserve the original object as opaque instead of rounding it into validity.

## Migration Plan

Additive support for previously rejected negative coordinates. Malformed or out-of-range coordinates now reject explicitly. Use the isolated checkout and publish only this change after focused native and documentation checks. No wire regeneration or full host suite is required.

Experimental distinction: an imported xfrm with long.MaxValue offset stays opaque in the native artifact and its no-op preserves the original bytes. PPJ projection independently rejects its frame above 100000 points. The malformed-source test checks both outcomes rather than asserting every invalid native frame is representable in PPJ.
