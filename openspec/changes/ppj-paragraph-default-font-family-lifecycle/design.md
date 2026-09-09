## Context

See proposal.md. Existing ApplyLatinFont owns one a:latin containing only typeface. Full default-style writes can rebuild unrelated effects; the recent scalar patch avoids that for bold/italic/size.

## Goals / Non-Goals

Complete direct Latin font presence for ordinary text/shape paragraphs with fixed topology. Script-font lifecycles, inherited placeholders and installed-font substitution remain separate.

## Decisions

Extend field masking, exact capability checking and the cloned default style with FontFamily presence. Extend the targeted default-style patch to call ApplyLatinFont only when family changes; keep other native children untouched. Reuse existing name validation (nonblank, at most 255 characters).

Authored BuildTextStyle also writes an East Asian fallback from fontFamily. The font-only source fixture explicitly removes that generated a:ea before projection; the mixed-font fixture keeps it and a:cs. Source-bound Latin edits do not propagate the authored fallback to other script fonts.

Reject extra attributes and children before editing a Latin font. Inspect serialized leaf content before typed reads, since SDK leaf ChildElements alone may hide unsupported XML. Do not broaden this field into font metadata reconstruction.

## Risks / Trade-offs

Unknown leaf content could be lost → test extra metadata and child-content rejection with unchanged source bytes. Token or family spelling could drift → verify freshly projected strings and preserve other script fonts, run fonts and unknown siblings. Shared whole-default-style baseline failure remains separately documented.
