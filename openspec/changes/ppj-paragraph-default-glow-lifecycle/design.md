## Context

See proposal.md. PPJ glow has color, required radius 0..1000pt and optional opacity/token. Default-run construction currently always writes alpha, and glow reading recognizes only glow/outer-shadow lists. Other default effects have similarly narrow whole-list readers, so removing glow can change which siblings appear in the native model.

## Goals / Non-Goals

Own one direct glow without changing sibling effect XML. Stabilize existing modeled default effects under glow edits; their independent edit authorities remain separate. Effect DAGs, duplicate lists/glows and unmodeled glow descendants remain source-owned.

## Decisions

Add exact glow authority and per-paragraph raw comparison. Construct paragraph glow with optional alpha presence, plain theme identity, transformed color resolution and existing radius/opacity limits. Source-bound declared grammar colors take precedence; other owners retain their existing construction behavior.

Isolate each unique direct default-effect node before calling its existing bounded reader. Validate raw descendants first so hidden alpha children and illegal text cannot become editable. This keeps sibling semantic fields stable without adopting a whole unknown graph or adding their edit authority.

Patch an existing glow in place; add a new glow using native child ordering. Deletion removes only that glow and an empty attribute-free effect-list wrapper. Preserve list attributes, known/unknown sibling nodes, direct runs and neighboring paragraphs. Whole default-style clearing removes glow only when its direct owner is proven.

## Risks / Trade-offs

Sibling effects appear only after glow deletion → isolated reading plus mixed-effect deletion/restoration evidence.

Source nested alpha data hidden by typed leaf access → inspect raw XML and retain rejection/preservation fixtures.

Theme/grammar collision or implicit alpha 1 changes source meaning → distinguish omission, explicit zero/one, RGB/RGBA, plain scheme tokens and grammar values in focused cases.

Existing whole-default-style baseline failure → keep its recorded boundary separate from this field's lifecycle.
