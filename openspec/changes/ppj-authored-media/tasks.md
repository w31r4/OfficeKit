## 1. Contract and model

- [x] 1.1 Add bounded PPJ media parsing and semantic rules for supported MIME, explicit poster, trim, loop, and mute.
- [x] 1.2 Add the additive `PresentationMedia` wire payload and regenerate bindings without changing protocol version 2.

## 2. Native compilation

- [x] 2.1 Add content-addressed media asset validation and NativeAOT PPJ lowering.
- [x] 2.2 Add canonical PowerPoint media picture, relationships, extension, and shared media-part writer.
- [x] 2.3 Extend the canonical timing writer for audio/video playback, trim, loop, mute, and coexistence with object animations.

## 3. Recovery and Agent surface

- [x] 3.1 Prove embedded PPJ and media/poster assets recover exactly while imported unknown media remains opaque.
- [x] 3.2 Update PPJ Help ownership, capability registry, generated `ppj.md`, focused media/layers guidance, and coverage.

## 4. Lean verification and integration

- [x] 4.1 Add one comprehensive authored media build/reimport/recovery test to the existing PPJ test surface.
- [x] 4.2 Run the PPJ contract, proto, NativeAOT, and Skill-maintainer narrow gates; record unverified host playback honestly.
- [x] 4.3 Commit the spec, runtime, guidance, and evidence atomically, then fast-forward the feature branch and coordinated main without force push.
