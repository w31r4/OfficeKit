# Proposal: allow typed grammar tokens for text language

Let the existing `text.language` field accept a declared `string` grammar
token in authored PPJ. The resolved value still has to pass the existing
bounded BCP-47 language-tag check before it is written to `a:rPr/@lang`.

This is a narrow K-05 expression improvement: it reuses the current wire and
native language owner, and does not infer locale, direction, font fallback,
or source-bound theme values.
