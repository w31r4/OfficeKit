---
name: artifact-template-evidence-ledger
description: "Create evidence-led presentations with an editorial visual grammar: full-bleed image fields, readable data, causal lines, and decisive conclusions. Use when the user wants a measured, research, strategy, or decision deck with a restrained visual system."
---

# Evidence Ledger

This is a visual reference template, not a page clone. Read this file before
authoring and inspect the representative images in `assets/examples/`.

## Design grammar

- Start with the communication claim, then choose the visual carrier that lets
  the audience verify it: chart, comparison, causal line, table, or image.
- Use a restrained ink / paper / cyan / ochre palette with clear role contrast;
  preserve the deck's own colors when adapting the grammar to new material.
- Use a strong typographic hierarchy, short labels, horizontal rules, and
  deliberate density changes between opening, evidence, and decision pages.
- Let one visual argument occupy the page. Supporting labels explain the
  relationship; decoration never competes with the evidence.
- A full-bleed image field may sit behind editable text. Keep foreground text
  legible and keep charts, lines, and labels unobstructed.

## Authoring

1. Form a deck-specific brief, narrative, and design grammar from the user's
   evidence; do not copy the example wording or fixed coordinates.
2. Choose a page archetype that matches the claim. Use native charts, tables,
   lines, shapes, and text for evidence; use an image only when it has a clear
   evidence, context, identity, or atmosphere role.
3. For a full-bleed static image behind editable content, use
   `slide.setNativeBackgroundImage({ blob, fit: "stretch" })`. Use
   `slide.setBackgroundImage(...)` when the image must remain a movable scene
   element. Inspect `slide.elements` before changing layer order.
4. Render the page, check text overflow, object overlap, contrast, crop, and
   source/credit requirements, then reopen the exported PPTX before delivery.

The examples are visual evidence only. They do not grant permission to copy
their text, page geometry, or assets into another presentation.
