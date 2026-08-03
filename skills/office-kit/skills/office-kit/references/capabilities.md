# Optional visual capabilities

Skills reason about two abstract capabilities supplied by the active agent or
host:

- `image_view`: the agent can understand a rendered image or preview.
- `image_generate`: the agent can create or modify an image asset.

Do not assume either capability. Use this route, in order:

```text
user or template asset
→ native Office shapes, charts, tables, and typography
→ optional generated asset
→ explicitly labelled placeholder
```

| `image_view` | `image_generate` | Route | Required report |
| --- | --- | --- | --- |
| yes | yes | Generate or adapt assets, then perform visual and structural QA. | `visualReview: "complete"` only after the rendered result was understood. |
| yes | no | Use user/template assets and native Office primitives. | Review the render and report structural findings. |
| no | yes | Use generated assets only for decorative or low-risk content. | Mark important imagery `requires-human`; do not claim visual approval. |
| no | no | Use templates, native shapes, charts, tables, and typography. | Run structural QA and report `visualReview: "unavailable"`; ask for an asset when imagery is essential. |

PowerPoint shapes, connectors, charts, and text are valid visual composition
tools. Image generation is optional, never a prerequisite for a coherent deck.
When there is no visual input, structural QA must still check file format and
dimensions, image count and placement/crop, text overflow, object overlap,
contrast values, and page/slide geometry. These checks are not an aesthetic
judgement.
