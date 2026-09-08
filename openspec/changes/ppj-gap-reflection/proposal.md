# Proposal: add authored reflection effects to PPJ

Add a bounded `reflection` field for authored shape styles and image styles so
PPJ can express the direct DrawingML `a:reflection` owner.

The profile owns blur, start/end opacity, distance, and direction. Native
lowering uses the standard full-span reflection positions. Imported effect
graphs, 3-D, and other unsupported effects remain opaque.
