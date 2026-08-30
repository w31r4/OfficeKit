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

Use an image as the page `background` when it is a true non-selectable slide
surface. PPJ compiles it into native `<p:bg>` image paint, not a backmost picture
shape. Native image backgrounds support `stretch`, `cover`, `contain`, explicit
signed crop, direct opacity, and parameter-free `tile`. Use an image element
when it must remain independently selectable, masked, bordered, shadowed,
animated, or reordered.

```json
{
  "background": {
    "type": "image",
    "asset": "wetland-aerial",
    "fit": "cover",
    "opacity": 0.78
  }
}
```

`cover` and `contain` lower deterministically from the declared asset dimensions
and slide frame. An explicit `crop` overrides that calculation. `tile` emits
only the portable default DrawingML tile profile; arbitrary tile scale, offset,
alignment, mirroring, external links, and extra image effects fail closed. An
imported native background can change only when the page `nativeRef` issues
`setBackground` for that exact source revision.

A page background may also use a native linear or centered radial gradient.
Choose it as a surface hierarchy or directional-light device, not as substitute
imagery. Keep the stop count small, maintain text contrast across the entire
field, and use a foreground image element when a photographic subject must be
cropped or animated independently.

Keep the subject, intended focal point, and important edges inside the crop.
Use `cover` for image-led regions, `contain` for diagrams or logos that must not
crop, `tile` only for a genuinely repeatable texture, and explicit crop values
when reproducibility matters. Re-render after font, crop, mask, or z-order
changes.

Shape and text-container fills can use the same bounded native image-paint
profile. This is useful when a photograph or texture must be clipped by editable
shape geometry:

```json
{
  "type": "shape",
  "id": "material-window",
  "frame": { "x": 72, "y": 110, "width": 360, "height": 250 },
  "geometry": { "kind": "preset", "preset": "arc" },
  "style": {
    "fill": {
      "type": "image",
      "asset": "material-photo",
      "fit": "cover",
      "opacity": 0.9
    }
  }
}
```

This creates a real `a:blipFill` on the shape. It is not a rasterized silhouette.
Imported canonical image fills project back into PPJ and may be changed only
when `nativeRef.capabilities` issues `setFill`.

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
