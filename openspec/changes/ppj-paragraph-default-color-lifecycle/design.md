## Context

See proposal.md. Default style already models RGB/scheme colors and optional alpha; projection uses RGB/RGBA strings or theme token objects. Color and gradient are mutually exclusive. Native scheme reading also tolerates source-owned luminance transforms, which the PPJ color field does not fully express.

## Goals / Non-Goals

Complete direct solid-color presence, including alpha zero/one, without touching unrelated paragraph/run paint. Gradient lifecycle, complex transforms and host typography remain separate.

## Decisions

Extend exact field authority and per-paragraph raw comparison. Use the existing color resolver for RGB and palette inputs, explicitly resolve declared grammar tokens first, and retain untransformed standard source theme tokens as scheme values.

Preserve direct alpha presence from RGBA strings or explicit alpha objects, including one; native alpha rounds at 1/100000 with existing away-from-zero conversion. RGB projection retains its existing eight-bit alpha representation; untouched source precision is never recomputed.

Patch only changed solid fill via the existing fill writer. Guard replacements and deletion of source luminance transforms, including the no-default-style deletion route. Unknown/no-fill/multiple paint stays source-owned; unrelated scalar edits retain it. Gradient/color conflicts and gradient changes require their own authority.

## Risks / Trade-offs

Opacity precision in neighboring paragraphs → compare their exact native XML during another paragraph's edit.
Theme/grammar collision → verify explicit grammar precedence and standard scheme preservation.
Luminance transforms silently removed → guard targeted writing and whole-wrapper deletion, then test rejection plus unrelated preservation.
Known whole-default-style baseline failure → retain the documented exclusion and separate host acceptance.
