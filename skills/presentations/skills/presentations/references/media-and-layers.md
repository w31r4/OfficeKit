# Media and layers

Images, SVG, video, and audio must carry evidence, identity, explanation, or
atmosphere. Do not add an asset merely to fill empty canvas.

## Asset contract

Every PPJ asset uses a relative URI, exact MIME type, SHA-256, rights status,
and accessibility information. Keep files beside the program or inside its
managed asset workspace. Do not place remote URLs, network instructions, or
large base64 payloads in PPJ.

Prefer sources in this order:

1. user-provided, brand, template, or official assets;
2. evidence images tied to the claim;
3. rights-compatible editorial or contextual imagery;
4. generated imagery when the host provides it and its status is recorded;
5. native vectors or no image when no suitable asset exists.

Do not present a generated or decorative image as documentary evidence. Keep
required attribution visible or in an explicit credits page according to the
license.

## Layer stack

`pages[].elements[]` is the true back-to-front z-order. A common image-led page
may use:

```text
background image
→ crop or color field
→ scrim/mask for contrast
→ evidence or identity layer
→ editable title and annotation
→ foreground action or source
```

Use an image as the page `background` only when its behavior belongs to the
slide background. Use an image element when it must remain independently
editable, crop-able, animated, or reordered. Masks, opacity, borders, and
shadows belong to the image element; overlays are ordinary shapes placed after
the image and before text.

Keep the subject, intended focal point, and important edges inside the crop.
Use `cover` for image-led regions, `contain` for diagrams or logos that must not
crop, and explicit crop values when reproducibility matters. Re-render after
font, crop, mask, or z-order changes.

## Protect content

Text contrast must survive the actual image, not an imagined average color.
Scrims should be only strong enough to establish hierarchy. Decorative fields
may overlap each other, but may not hide text, sources, chart evidence, product
details, or the image subject.

Avoid repeated stock photographs, identical hero crops, gratuitous shadows,
full-page texture that reduces legibility, and photo mosaics without a narrative
relationship. A page can be image-free when data, geometry, or typography is
the stronger carrier.
