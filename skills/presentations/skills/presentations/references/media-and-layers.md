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

## Authored audio and video

Use a `media` element only for a local embedded asset whose playback role is
part of the presentation. PPJ source-free authoring supports MP4 video and
MP3, M4A, or WAV audio. Every media element requires an explicit image
`posterAsset`; the poster is the editable native picture surface seen before
playback and in static rendering.

```json
{
  "id": "field-observation",
  "type": "media",
  "name": "Field observation video",
  "role": "primary evidence",
  "frame": { "x": 528, "y": 96, "width": 360, "height": 240 },
  "accessibility": {
    "decorative": false,
    "description": "A short field observation showing the measured condition."
  },
  "mediaType": "video",
  "asset": "field-observation-mp4",
  "posterAsset": "field-observation-poster",
  "startAtMs": 1200,
  "endAtMs": 400,
  "loop": false,
  "mute": true
}
```

`startAtMs` and `endAtMs` are bounded leading and trailing trim offsets, not
timeline expressions. `loop` and `mute` compile to native playback state.
OfficeKit owns the media relationships, click action, canonical timing nodes,
poster relationship, and package part names; PPJ owns only the typed state
above. Keep media below 64 MiB per asset and within the deck's aggregate asset
budget.

Static render and structural review prove the poster and package graph, not
playback. Record desktop evidence separately when actual playback matters.
Third-party media timing remains opaque/source-bound: importing it does not
authorize rewriting triggers, bookmarks, effects, or an unfamiliar timing
graph.

## Imported paired SVG pictures

Some PowerPoint pictures carry two assets: a PNG or JPEG compatibility image
in `image.asset` and an SVG for modern hosts in `image.svgAsset`. Treat them as
different roles. Do not replace the fallback merely because the vector artwork
is the intended edit.

When inspection issues `replaceSvg` for `image.svgAsset`:

1. create a new local `image/svg+xml` asset declaration with exact bytes and
   SHA-256;
2. point only `svgAsset` at the new declaration;
3. keep `asset`, the element ID, frame, crop, nativeRef and array position;
4. build, re-import and render the edited page.

The compiler replaces only the proven SVG relationship. It does not create or
remove a fallback pair and does not rasterize the new SVG. An unchanged raster
fallback is intentional compatibility state, so legacy-host appearance must be
checked separately when it matters. A standalone authored SVG remains an
ordinary `image.asset`; do not add `svgAsset` to source-free PPJ.

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

A native solid background also retains direct opacity. Use either an
alpha-bearing color or an explicit `opacity`; the explicit value wins:

```json
{
  "background": {
    "type": "solid",
    "color": "#0A84FF",
    "opacity": 0.45
  }
}
```

The page owns one background paint, not a stack of background paints. To place
a color scrim over a background photograph, keep the photograph as the native
image background and put an ordinary translucent shape first in
`elements[]`, before foreground text.

### Element visibility and edit locks

All typed elements can use `hidden` and `locked`. Use `hidden: true` to keep a
stable object in the program while excluding it from the visible slide. Use
`locked: true` after a background, guide, media poster, or finished composition
is placed correctly and should resist accidental selection or movement.

These fields do not change layer order: `elements[]` remains the only
back-to-front stack. They also do not hide the page or protect the file from a
user with edit access. On an imported deck, change a field only when the
object's `nativeRef.capabilities` includes the matching `setHidden` or
`setLocked`; unfamiliar partial lock combinations remain source-owned.

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

Use an image element when it must remain independently selectable, cropped,
animated, or reordered. Preset masks, direct opacity, bounded RGB borders, and
one bounded outer shadow belong to that image element; overlays are ordinary
shapes placed after the image and before text. On imported files, change only
the effect values and mask adjustments returned by inspection. Unfamiliar
effect graphs and custom mask topology stay source-bound instead of being
flattened.

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
topology remain fixed.

Use a custom mask only when the silhouette carries a real editorial or brand
role that no preset geometry expresses. It uses the same finite literal-path
vocabulary as a custom shape:

```json
{
  "type": "image",
  "id": "field-photo",
  "frame": { "x": 552, "y": 72, "width": 360, "height": 396 },
  "asset": "field-photo",
  "fit": "cover",
  "mask": {
    "kind": "custom",
    "viewBox": { "x": 0, "y": 0, "width": 160, "height": 160 },
    "paths": [{
      "fill": true,
      "stroke": false,
      "commands": [
        { "op": "moveTo", "x": 20, "y": 0 },
        { "op": "lineTo", "x": 160, "y": 0 },
        { "op": "lineTo", "x": 140, "y": 160 },
        { "op": "lineTo", "x": 0, "y": 120 },
        { "op": "close" }
      ]
    }]
  }
}
```

The compiler writes a native picture `<a:custGeom>`; it does not rasterize the
mask. Keep the path count and command count small, and do not invent irregular
blobs merely to decorate empty space. Canonical imported literal masks can be
inspected and projected into PPJ, but their path topology remains source-owned:
without a separately issued mutation capability, changing those paths fails
before output. Guide formulas, handles, connection sites, text rectangles and
other richer custom-geometry graphs remain opaque-preserved.

## Protect content

Text contrast must survive the actual image, not an imagined average color.
Scrims should be only strong enough to establish hierarchy. Decorative fields
may overlap each other, but may not hide text, sources, chart evidence, product
details, or the image subject.

Avoid repeated stock photographs, identical hero crops, gratuitous shadows,
full-page texture that reduces legibility, and photo mosaics without a narrative
relationship. A page can be image-free when data, geometry, or typography is
the stronger carrier.
