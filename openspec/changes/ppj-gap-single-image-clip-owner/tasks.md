## 1. Contract and lowering

- [x] 1.1 Add the bounded single-image clip predicate to PPJ semantic validation and verify unsupported clip forms receive `ppj.compositing.clipUnsupported`
- [x] 1.2 Lower a valid clip's preset and adjustments to the existing `PresentationImage` mask owner, with a defensive compiler check and no wire change

## 2. Evidence and documentation

- [x] 2.1 Add a focused authored compile/XML/snapshot-free projection test covering mask, opacity coexistence, adjustment recovery, and unsupported topology
- [x] 2.2 Update the PPJ reference, coverage matrix, and K-04 backlog with the exact bounded boundary and evidence path

## 3. Verification

- [x] 3.1 Run the focused codec tests, OpenSpec strict validation, `git diff --check`, and the apply status check; record the actual result in the handoff
