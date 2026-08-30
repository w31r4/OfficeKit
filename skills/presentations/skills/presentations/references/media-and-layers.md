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

The authored native-background profile is intentionally narrow: opaque solid
color, bounded gradient, or an opaque image with `fit: "stretch"`. A cropped,
contained, tiled, or translucent picture must be the first image element in
`pages[].elements[]`; place any scrim and editable foreground content after it.
Unsupported background paint fails before PPTX output instead of silently
changing the crop or alpha.

A page background may also use a native linear or centered radial gradient.
Choose it as a surface hierarchy or directional-light device, not as substitute
imagery. Keep the stop count small, maintain text contrast across the entire
field, and use a foreground image element when a photographic subject must be
cropped or animated independently.

Keep the subject, intended focal point, and important edges inside the crop.
Use `cover` for image-led regions, `contain` for diagrams or logos that must not
crop, and explicit crop values when reproducibility matters. Re-render after
font, crop, mask, or z-order changes.

## Image masks

A preset image mask is native editable geometry, not a rasterized cutout. It
uses the same finite profile table as a shape:

```json
{
  "type": "image",
  "id": "portrait",
  "frame": { "x": 612, "y": 92, "width": 252, "height": 336 },
  "asset": "portrait-photo",
  "fit": "cover",
  "mask": {
    "kind": "preset",
    "preset": "round2SameRect",
    "adjustments": [18000, 6000]
  }
}
```

The adjustment array is complete or omitted for native defaults. Its order and
defaults come from the [generated PPJ preset table](ppj.md#preset-geometry-adjustments).
An imported picture can change only these values when `nativeRef.capabilities`
issues `setImageMask`; the preset identity, crop, asset, frame, and native
topology remain fixed. Custom-path picture masks remain source-preserved and
fail closed for authored or source-bound mutation.

## Protect content

Text contrast must survive the actual image, not an imagined average color.
Scrims should be only strong enough to establish hierarchy. Decorative fields
may overlap each other, but may not hide text, sources, chart evidence, product
details, or the image subject.

Avoid repeated stock photographs, identical hero crops, gratuitous shadows,
full-page texture that reduces legibility, and photo mosaics without a narrative
relationship. A page can be image-free when data, geometry, or typography is
the stronger carrier.
