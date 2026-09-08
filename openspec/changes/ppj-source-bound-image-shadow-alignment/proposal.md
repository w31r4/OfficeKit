# Proposal: add source-bound picture shadow alignment ownership to PPJ

Expose a bounded `image.shadow.alignment` owner for an existing imported
picture outer shadow. The owner issues one native leaf so a caller can change
the existing direct `a:outerShdw/@algn` token without rewriting the picture
payload, mask, border, or other shadow fields.

This slice does not infer a missing geometry attribute or broaden picture
effects. Only a strict direct `a:effectLst/a:outerShdw` with one explicit
DrawingML alignment token from the canonical nine-value set is editable;
missing values, additional effects, unknown children, extensions, malformed
topology, and other owners remain source-owned or fail closed.
