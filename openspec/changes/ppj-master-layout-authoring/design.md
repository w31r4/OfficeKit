# Design

## Language

Master and layout definitions live under `design` because they are persistent
presentation design state, not reusable page components:

```json
{
  "design": {
    "masters": [{
      "id": "master-main",
      "name": "Main master",
      "background": { "type": "solid", "color": { "token": "paper" } },
      "textStyles": {
        "title": [{
          "level": 0,
          "defaultText": { "font": "display", "size": 28 }
        }]
      },
      "placeholders": []
    }],
    "layouts": [{
      "id": "layout-title",
      "name": "Title",
      "master": "master-main",
      "layoutType": "title",
      "placeholders": [{
        "id": "layout-title-heading",
        "name": "Title",
        "placeholderType": "title",
        "index": 1,
        "frame": { "x": 48, "y": 42, "width": 624, "height": 72 }
      }]
    }]
  },
  "pages": [{
    "id": "page-1",
    "role": "opening claim",
    "layout": "layout-title",
    "elements": []
  }]
}
```

Array order is native master/layout order. Source-free compilation accepts one
master and at most 256 layouts. Master text styles contain at most one paragraph
style for each level zero through eight in each of `title`, `body`, and `other`.
Owner-local placeholders use explicit point frames and the closed text types
`title`, `body`, `centered-title`, and `subtitle`.

## Native lowering

The authored C# compiler maps PPJ definitions onto the existing typed
`PresentationMaster`, `PresentationLayout`, `PresentationMasterTextStyles`, and
`PresentationPlaceholder` wire messages. It does not write OOXML itself. The
existing PPTX writer independently revalidates one-master topology, layout type,
placeholder type, direct frame, background assets, and slide bindings.

When no master/layout state is declared, the existing canonical fallback master
and blank layout remain unchanged. When state is declared, pages must name a
declared layout and page-local placeholders must explicitly match a master or
layout placeholder by native type and index.

## Source projection

Third-party PPJ records the source slide's stable layout identity in
`pages[].layout`. It does not synthesize editable master/layout definitions from
an arbitrary native graph. Source-bound compilation requires the projected
layout value to remain unchanged, while the original PPTX remains the authority
for Master/Layout parts and relationships.

## Maintenance proof

The capability registry gains an API-to-PPJ-path table for every `ppj-state`
Help API. The maintainer validates exact key parity and that each referenced
root path belongs to the schema. This turns the registry from a category label
into a discoverable mapping contract.
