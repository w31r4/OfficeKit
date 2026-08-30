## Context

The public Presentation runtime already exposes `presentation.reuseSourceSlide`
after `slide.cloneCapability` proves the concrete OPC ownership graph. The
NativeAOT writer accepts a `PresentationSlide` with `clone_source`, recursively
copies safe owned parts, shares only proven immutable or identity resources,
and validates the resulting package. PPJ projection intentionally withheld
that capability because the language had no state representation for a
pending native clone.

## Goals / Non-Goals

**Goals:**

- Express exact source slide reuse as finite declarative PPJ state.
- Reuse the existing native clone graph proof and writer.
- Keep the pending clone immutable until build/reimport.
- Preserve the original source page and all unrelated package content.

**Non-Goals:**

- Adding a procedural `clone()` command list.
- Copying arbitrary PPJ-authored pages through `sourceClone`.
- Editing a pending clone before its first build/reimport.
- Reusing one source page more than once in the same compile.
- Combining the pending clone with section/custom-show edits.
- Exposing part paths, relationship IDs, or raw OOXML.

## Decisions

### 1. `sourceClone` is a bounded page macro

A new page may declare:

```json
{
  "id": "page-summary-copy",
  "role": "source continuation",
  "elements": [],
  "sourceClone": {
    "page": "page-summary",
    "capability": "cap_..."
  }
}
```

The referenced source page remains in `pages`. The clone must immediately
follow it, carries no nativeRef, and has an empty explicit element array. Its
visible content is the exact finite expansion of the proven source graph, just
as a component instance expands from a definition. The new page ID is PPJ
program state; the capability ID binds the request to the fresh projection.

### 2. The source page grants authority

Projection advertises `duplicate/pageClone` only when
`slide.source.clone_capability.supported` is true. The compiler reprojects the
exact source package, resolves the referenced page and capability, and then
sets the native wire `clone_source`. A copied or invented descriptor therefore
cannot grant itself authority.

### 3. Pending clones are deliberately immutable

The clone page may contain only its ID, role, optional claim, empty elements,
and `sourceClone`. Name, layout, background, notes, visibility, transition,
animation, and element changes require a normal source-bound projection after
the first build. This matches the mature runtime boundary and prevents PPJ
from pretending it can safely edit an object graph that does not exist yet.

### 4. Routing state stays unchanged

Sections, custom shows, and presentation comments remain byte-equivalent to
the baseline during the pending clone build. The native clone stays outside
custom shows and follows the source page in presentation order. After
reimport, ordinary capability-issued route edits may include the new page.

## Risks / Trade-offs

- [The page looks incomplete in JSON] -> Document `sourceClone` as an explicit
  finite macro whose content is inspected on the referenced page.
- [Stale capability replay] -> Resolve both page and capability against a fresh
  projection of the exact source SHA-256.
- [Agent edits the clone immediately] -> Reject every explicit native field or
  element until build/reimport.
- [Clone changes section/show topology] -> Require unchanged route state in
  the same compile.
- [Repeated clone creates ambiguous ownership] -> Permit one pending clone per
  source page, matching the existing native writer.

## Migration Plan

Additive optional page state and capability vocabulary. Existing authored and
source-bound PPJ remains valid. Source-free PPJ cannot use `sourceClone`.

## Open Questions

None.
