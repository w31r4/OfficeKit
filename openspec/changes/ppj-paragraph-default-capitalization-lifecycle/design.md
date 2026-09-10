## Context

See proposal.md. Authored/default-run readers already use FontCaps with none/small/all. The native cap enum has the same tokens; unknown tokens are omitted from modeled style but remain in source XML.

## Goals / Non-Goals

Complete direct capitalization presence for ordinary text/shape paragraphs without rewriting text content. Run overrides, table defaults, inherited placeholders and host glyph rendering retain separate profiles.

## Decisions

Extend exact field masks/authority and the cloned default style. Reuse enum validation. Patch only changed cap in the scalar writer instead of rebuilding unrelated font/effect XML; assign null directly for deletion. Preserve explicit none separately from omission.

Reject replacement of unmodeled native cap and clear only modeled capitalization during default-style cleanup. Existing export validation allows retained source warnings when no new warnings are introduced. Preserve invalid native cap during no-op and unrelated assignment/removal; do not discard it to normalize the source. Replacement through the new capitalization field rejects.

## Risks / Trade-offs

None may collapse into omission → assert native/fresh PPJ presence. Capitalization may be implemented as text rewriting → compare all run text/native XML. Unknown cap may disappear during unrelated deletion → a focused invalid-source fixture requires no-op identity and replacement rejection and preservation during unrelated assignment/removal. Host typography remains unverified.
