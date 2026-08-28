# Layered composition

Use this reference when a page depends on cross-type overlap, an image-led
background, or z-order changes in an imported PPTX.

## Think in one scene stack

Every direct slide object shares one bottom-to-top stack:

```text
background field or image
→ contrast treatment
→ evidence and structure
→ editable copy and labels
→ foreground emphasis
```

`slide.elements.items` exposes that order. Shapes, textboxes, images, tables,
charts, connectors, and groups also remain available through their typed
collections, but those collections do not define export order.

Use geometry only when it communicates contour, scale, position, connection,
or enclosure. An overlapping layer must have a named information job. Do not
add rings, blobs, panels, arrows, or scrims merely to fill space.

## Public surface

```js
const layers = slide.elements.items;
const index = element.stackIndex;
const capability = element.zOrderCapability;

element.sendToBack();
element.bringToFront();
element.moveBefore(peer);
element.moveAfter(peer);

const inspected = presentation.inspect({ kind: "layer" });
```

`moveBefore` places an element immediately behind a peer; `moveAfter` places it
immediately in front. Both elements must belong to the same slide or group
stack. Reinspect after reopening an exported file; object IDs and capabilities
are revision-bound evidence, not permanent selectors.

## Image-led pages

Use a full-slide image when the image itself carries context, evidence,
identity, or emotion. Preserve an important subject and information-bearing
region when cropping. Add a scrim only when it makes foreground content
readable without destroying the image's evidence.

There are two deliberate image routes. Use `slide.setNativeBackgroundImage(...)`
when the image is a true full-bleed backdrop that must sit below every slide
object:

```js
slide.setNativeBackgroundImage({
  blob: await FileBlob.load(asset.path, { type: asset.mimeType }),
});
```

This writes native `p:bg/p:bgPr/a:blipFill` and supports one embedded image
with `fit: "stretch"`. It is not a scene-stack element, so it cannot be
reordered or animated. Use `slide.setBackgroundImage(...)` below when the
image needs crop, cover/contain behavior, z-order changes, or animation; that
route writes an ordinary editable picture layer.

```js
const photo = slide.setBackgroundImage({
  blob: await FileBlob.load(asset.path, { type: asset.mimeType }),
  fit: "cover",
  accessibility: { description: "Industrial heat exchanger at dusk" },
});

const scrim = slide.shapes.add({
  name: "contrast field",
  geometry: "rect",
  position: { ...slide.frame },
  fill: { color: "#071A24", opacity: 0.62 },
  line: { style: "none" },
});

scrim.moveAfter(photo);
headline.bringToFront();
```

Repeated `setBackgroundImage` calls replace the authored background image.
`clearBackgroundImage()` removes it. This is an editable image layer, not a
rasterized page.

Do not assume one global opacity. Render the actual crop with the intended type
color and choose the lightest treatment that preserves legibility. A local
gradient, side field, or text shadow may be better than darkening the whole
image when the composition supports it.

## Imported slides

Preserve the source's native order by default. Before reordering:

1. inspect `kind: "layer"`;
2. resolve the exact target and peer;
3. require `zOrderCapability.editable === true`;
4. apply only the declared order change;
5. export, reopen, reinspect, and render the affected page;
6. compare the package footprint and non-target pages.

Only capability-proven direct imported elements may move. Nested imported group
children and unknown SlidePart topology remain source-bound. An authored
overlay on an imported slide must stay above the complete source-bound prefix;
OfficeKit will not insert a new background image beneath preserved native
content. Use a source-derived slide or preserve/reuse an existing source image
instead.

Do not combine a source-bound reorder with an unrelated authored overlay,
deletion, or broad rewrite in the same export. Reopen the reviewed revision
before the next mutation.

## Review the rendered relationship

Inspect both the stack and the rendered page. Layout boxes alone cannot prove
that a translucent or crossing object is readable.

- A chart line, marker, label, axis, connector, arrowhead, or causal route must
  not be hidden by an opaque object or by a label placed on top of the evidence.
- Do not separate real series merely to avoid overlap when the overlap carries
  meaning. Use transparency, direct-label offsets, local masks, or a different
  valid chart form while preserving the data relationship.
- Confirm that foreground text remains editable, legible, and visually tied to
  the image or evidence beneath it.
- Confirm crop, resolution, contrast, attribution, alt text, and source rights.
- Report `visualReview: "unavailable"` when no Agent or human understood the
  render. Structural stack inspection is not visual approval.
