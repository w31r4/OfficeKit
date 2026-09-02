---
name: "Presentations — Shared What / What-kind / How (experiment)"
description: "Create or edit editable PowerPoint presentations through PPJ using an explicit three-layer communication, scenario, and construction route."
---

# Shared What / What-kind / How

This is an experimental route. It uses the production PPJ capability pack and
the shared invariants at [common invariants](../../common/invariants.md).

## Route

Follow this order and record the decisions in the PPJ `intent` and `design`:

1. **What** — read [communication contract](../../common/what.md). Identify
   audience, job, expected change, evidence boundary, delivery mode, and page
   responsibilities.
2. **What-kind** — read [route and scenario contract](../../common/what-kind.md).
   Choose one scenario, one design source, and a delivery mode. Read exactly
   the selected guide in `../../common/references/scenarios/`.
3. **How** — read [construction contract](../../common/how.md). Write the
   deck-specific Design Grammar and attention contract, then load only the
   carrier-specific PPJ references.

## Create

Write one strict `.ppj` program. Use typed elements and stable IDs; never use
MJS/JSX, raw OOXML, XPath, relationship IDs, or a second authoring engine.
Choose a primary carrier from the declared relationship. Do not fill empty
space with cards, random geometry, icons, or stock images. If a page contains a
chart, line, table, mask, background image, or label, protect its evidence and
check the true back-to-front order.

If an image is required, read `../../common/references/image-sourcing.md` and
use the shared `officekit image` flow. The Agent chooses the query and crop;
record the asset hash, rights, source, and alt text.

Build and review in this order:

```text
check → build → render → inspect occupancy/reading order/occlusion
→ repair the responsible layer → review → deliver
```

## Import and edit

For existing PPTX, use `officekit ppj import`, then read
`../../common/references/imported-native-ref.md`. Preserve source binding,
stable IDs, opaque objects, and non-target parts. Perform local semantic edits
before visual/delivery edits. Unsupported or stale mutations fail closed.

## Delivery

Return absolute PPJ/PPTX paths, hashes, review status, and honest structural,
render, Keynote, or PowerPoint evidence. This experimental arm does not change
the production default and does not treat a clean XML check as playback proof.
