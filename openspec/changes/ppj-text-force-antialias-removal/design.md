## Context

See proposal.md. force_anti_alias is optional field 33. The anchorCenter deletion path establishes validation, normalization and source-edit inference for optional booleans.

## Goals / Non-Goals

Complete direct hint presence with existing text topology and authority. This does not implement a renderer-specific anti-aliasing algorithm.

## Decisions

Add optional no_force_anti_alias at unused field 41. Selected false and setter/deleter conflicts are invalid. Preserve old setter encoding. Apply deletion to forceAA and normalize the marker away before semantic comparison. Extend shared boolean and wire tests, preserving anchorCenter/verticalAlignment. Whole-style rejection uses spaceFirstLastParagraph after forceAntiAlias becomes removable.

## Risks / Trade-offs

Old codecs ignore new fields → document updated codec requirement; PPJ emits this operation inside its updated compiler. Shared registry has concurrent preview edits → stage only owned reason changes. True/false mistaken for rendering proof → keep preview limitation and native-attribute evidence distinct.
