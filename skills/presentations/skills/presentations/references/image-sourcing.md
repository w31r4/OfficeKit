# Image sourcing

Load this reference only when the page plan names an image or icon role that
supplied, template, or official assets cannot fill. OfficeKit provides
deterministic search, acquisition, provenance, and audit primitives. The Agent
still decides whether a candidate serves the page.

## Choose the media role first

An image must provide evidence, identity, explanation, or atmosphere. Do not
search merely to fill empty canvas. Prefer, in order:

1. user-provided and template assets;
2. official assets or an approved brand system;
3. a Lucide icon for a small structural symbol;
4. an Openverse or Wikimedia photo/illustration;
5. an optional host-provided generated image;
6. an editable chart, diagram, native visual, or no image.

Use English search terms and preserve names, products, places, and other proper
nouns. Search for the subject and visual role, not the whole slide sentence.

## Search and select

```bash
officekit image search "institutional bitcoin trading" \
  --task <task-id> --kind photo --purpose evidence \
  --orientation landscape --max 5 --json
```

`search` returns candidates with `selectionMade: false`. Inspect three to six
plausible candidates when available. Judge subject relevance, crop potential,
visual coherence, resolution, rights, and credit burden. The CLI never selects
for the Agent and never hides a provider failure.

Ask the user only when the choice changes brand identity, uses a recognizable
person, carries material legal risk, conflicts with the intended direction, or
has no compliant candidate. Otherwise choose the strongest compliant fit.

When visual understanding is unavailable, deterministic Lucide icons may be
used for small structural roles. Do not select a decorative photo and claim a
visual review. Use a native visual, omit the image, or mark the crop for human
review.

## Register before use

Add a searched candidate:

```bash
officekit image add --task <task-id> --candidate <candidate-ref> --json
```

Register a supplied file:

```bash
officekit image add --task <task-id> --file ./assets/hero.png \
  --rights user-provided --json
```

Register a deliberate HTTPS source only when the source page and rights are
known:

```bash
officekit image add --task <task-id> \
  --url https://press.example.com/assets/product.png \
  --source-page https://press.example.com/media \
  --rights official-press-kit --json
```

CC BY additionally requires `--author` and `--license-url`. Allowed sourced
rights are Public Domain, CC0, CC BY, Lucide ISC, explicit permission,
user-provided/generated assets, and official press kits. Reject ShareAlike,
NonCommercial, NoDerivatives, and unknown rights. Provider metadata is evidence
of a declaration, not a legal guarantee.

The returned task path is content-addressed and read-only. Use it rather than a
remote URL or copied base64:

```js
const asset = /* JSON result from officekit image add */;
slide.images.add({
  blob: await FileBlob.load(asset.path, { type: asset.mimeType }),
  position: { x: 7.2, y: 0.7, width: 5.4, height: 6.1 },
  fit: "cover",
  accessibility: { description: "Concise description of the visible subject" },
});
```

For imported decks, use the inspected source-bound replacement capability
after registration. Keep the task asset receipt even when the target picture
retains its original geometry and crop.

## Resume and audit

```bash
officekit image list --task <task-id> --json
officekit image audit candidate.pptx --task <task-id> \
  --sources-output candidate.pptx.sources.json --json
```

`list` recovers registered assets and prior searches in a new context. `audit`
matches task receipts to actual PPTX media bytes and reports used, unused, and
unregistered images plus visible credit obligations.

Follow [Review and deliver](../tasks/review-deliver.md) for crop, contrast,
resolution, repetition, alt text, visible attribution, sources sidecar, and
visual-review status. Do not treat a clean rights audit as proof that the image
looks good or that a provider independently verified the license.
