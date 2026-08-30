## Why

PPJ already owns typed animation state and the NativeAOT codec already reads,
validates, writes, and reimports OfficeKit's canonical PowerPoint timing graph.
Imported pages even project recognized animations. The source-bound compiler,
however, treats `pages[].animations` as immutable, so the public language hides
the mature ability that formerly sat behind `slide.animations.add/remove`.

## What Changes

- Add `setAnimations` to the closed native capability vocabulary.
- Issue `setAnimations/animations` on imported pages whose native timing graph
  is canonical and editable, or whose graph is absent and safely addable.
- Do not issue the capability for opaque timing or a page participating in
  Morph; those graphs remain unchanged.
- Lower the complete requested `animations[]` state through the existing PPJ
  authored animation lowerer and `PptxTimingCodec` writer.
- Resolve PPJ stable target IDs back to their exact native element IDs,
  including supported descendants inside groups.
- Require build/reimport before an authored topmost overlay can be animated.

## Capabilities

### New Capabilities

- `ppj-source-motion-parity`: Capability-issued canonical object animation
  editing on imported pages.

### Modified Capabilities

None. The PPJ schema ID and Office wire protocol version remain unchanged.

## Impact

PPJ native capability schema, projection, source-bound lowering, one shared
animation lowerer, generated Skill guidance, coverage, and one existing
focused C# contract are affected. The timing XML codec, JavaScript runtime,
Office wire, transition/morph model, and public PPJ animation schema are reused.
