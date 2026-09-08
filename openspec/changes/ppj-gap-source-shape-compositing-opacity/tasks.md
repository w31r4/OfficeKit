## 1. Contract and projection

- [x] 1.1 Reuse the existing `compositing.opacity` schema and add a bounded
  source-shape owner rule to the native projector.
- [x] 1.2 Advertise `setOpacity/compositing.opacity` only for shapes that pass
  the owner-alpha proof.

## 2. Source-bound lowering and evidence

- [x] 2.1 Add absolute compound-alpha lowering for direct fill, gradient,
  image, outline, and shadow owners.
- [x] 2.2 Add the focused source-bound projection/edit/XML/reprojection test,
  including fail-closed mixed-alpha coverage.

## 3. Bookkeeping and gates

- [x] 3.1 Update coverage and the K-04/F-05 backlog wording.
- [x] 3.2 Run strict OpenSpec validation, native build/focused test, reference
  checks, and repository diff checks.
