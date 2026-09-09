## Context

See proposal.md. Native CustomGuides already stores ordered name/formula pairs and validates 17 operators, names, arity, prior-reference ordering and finite bounded evaluation. The text rectangle resolver shares this graph. PPJ currently rejects all custom-guide owners during shape projection.

## Goals / Non-Goals

Goals: expose that ordered guide graph for otherwise literal custom shape paths, with source add/edit/delete and reference preservation. Non-goals: custom adjustments, handles/sites, reference-backed path coordinates, masks/clips or host-exact layout.

## Decisions

- Use geometry.guides as a maximum-1024 list of name/formula pairs, reusing the existing formula grammar and native units. Formula whitespace may normalize on native read; no PPJ arithmetic evaluator is introduced.
- Empty or omitted list means no user guides and projects as omission. OfficeKit's private numeric-text-rectangle guide tail stays hidden by the existing native codec.
- Allow CustomGuides in the otherwise unchanged literal-path projection/edit profile. Keep adjustments/sites/handles and nonliteral paths guarded.
- Source setGeometry issues geometry.guides authority. Rebuild the requested graph together with unchanged paths/rectangle; validate the final graph so removing referenced guides fails unless references are edited in the same request.
- Rename the shared authored shape-only switch to allowShapeGraph; actual shapes accept guides/rectangle, masks and clips reject them rather than dropping state.
- Test one graph feeding a text rectangle; check authored/fresh identity, a formula edit's evaluated result, safe add/removal, dependency failures and non-target ZIP preservation. Preview limitations remain explicit.

## Risks / Trade-offs

- Formula strings mistaken for points -> retain DrawingML formula units and documented built-ins.
- Stale reference after graph change -> native final-graph validation.
- Private scaling guides exposed -> rely on the existing ordered private-tail decoding and test mixed rectangles.

## Migration Plan

Additive schema and source capability fields. No wire change; refresh source projections to obtain the new authority and regenerate documentation.
