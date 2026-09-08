## 1. Contract and lowering

- [x] 1.1 Extend the bounded image clip predicate to accept only validated custom geometry and verify malformed/custom-owner conflicts report `ppj.compositing.clipUnsupported`
- [x] 1.2 Lower one accepted custom clip to `PresentationImage.CustomMaskPaths` with a defensive compiler gate and no wire change

## 2. Evidence and documentation

- [x] 2.1 Add a focused authored compile/XML/snapshot-free projection test for custom path recovery and unsupported topology
- [x] 2.2 Update the PPJ reference, coverage matrix, and K-04/F-05 backlog with the exact custom-path boundary and evidence path

## 3. Verification

- [x] 3.1 Run the focused codec tests, OpenSpec strict validation, `git diff --check`, and the apply status check; record the actual result in the handoff
