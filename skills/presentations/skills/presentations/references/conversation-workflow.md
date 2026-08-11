# Conversational deck workflow

Use this for a net-new deck or broad redesign unless the user asks for one-pass
final delivery. Skip it for read-only work and a narrow existing-deck edit;
complete those under the ordinary inspect/edit/verify rules.

## 1. Clarify only material uncertainty

Infer audience, job, outcome, takeaway, sources, timing, and template state.
Ask only when uncertainty materially changes narrative, evidence, or visuals;
ask at most three questions in one turn. Choose safe routine aesthetics.

## 2. Build a working draft

- reopen and check facts, sequence, sources, package, and every slide;
- render every slide and fix deterministic clipping, overflow, overlap, broken
  connectors, and unresolved placeholders;
- record the honest visual-review status.

Write the working PPTX to a collision-safe path under `taskRoot`; keep every
input immutable. Give the user its absolute path and SHA-256, but label it as a
working draft. Do not call `ctx.publish` or describe the work as delivered.

Give a one-screen draft guide in the user's language:

- **Goal:** one sentence for what the audience should understand or do.
- **Structure:** three to six short section beats, never a slide dump.
- **Confirm:** at most three unresolved claims, assumptions, or high-impact
  design choices. Omit this section when nothing material is unresolved.
- **Try saying:** at most three concrete natural-language revision requests.
- **Draft:** the working PPTX path and its review status.

Prioritize impact, collapse non-blocking QA, and expand slides only on request
or when one needs a decision.

## 3. Revise through conversation

- For a local change, resolve the current slide/object again, edit only that
  scope, rerender it, and confirm unrelated slides remain stable.
- For a global change, edit the declared sections and recheck the narrative;
  state scope first only when materially ambiguous.
- After a timeout, restart, or stale hash, reimport and inspect the latest
  draft. Never reuse a stale locator.

Return actual changes, the top unresolved item, and the latest draft path.
Repeat the full guide only after broad redesign or on request.

## 4. Finalize explicitly

Explicit “finalize”, “deliver”, “定稿”, “交付”, or “可以了” permits complete
post-edit review and final publication. An explicit one-pass-final request is
acceptance in advance. Silence or the absence of further edits is not.

Reopen the latest draft; complete semantic, structural, render, optional text,
visual/human, accessibility, and delivery checks; then call `ctx.publish`.
Return its path, kind, SHA-256, status, and material limitations.
