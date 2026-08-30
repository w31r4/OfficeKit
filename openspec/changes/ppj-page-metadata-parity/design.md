## Context

The PPTX codec treats the direct common-slide-data name as bounded metadata and
already proves a requested name against the source SlidePart hash. Visibility is
owned by `PptxSlideVisibilityCodec`, which recognizes absence and the four XML
Schema boolean lexical forms, records a semantic hash, and fails closed for an
irregular root. PPJ currently stops before both paths.

## Goals / Non-Goals

**Goals:**

- Make imported `pages[].name` and `pages[].hidden` genuinely editable where the
  native codec already proves them.
- Keep capability issuance explicit and source-revision-bound.
- Change only the target SlidePart and recover exact state after reimport.

**Non-Goals:**

- Page reorder, section/custom-show membership, layout binding, page deletion,
  show settings, or presentation playback control.
- Repairing malformed native visibility values.

## Decisions

### 1. Name and hidden remain ordinary page fields

No operation list or edit-plan syntax is added to the PPJ program. The Agent
edits `name` and `hidden`; `nativeRef` explains whether the source page can
accept that state change.

### 2. Name uses the page source hash

Every valid imported page can issue `setName`. The page nativeRef object hash is
the exact SlidePart hash already re-proven by the codec before it writes the
direct `p:cSld/@name` value. Omission means an empty native name and is therefore
a supported clear operation.

### 3. Hidden requires visibility evidence

`setHidden` is issued only when the native root visibility is canonical. An
editable imported page always projects an explicit boolean. Removing that field
is rejected; the Agent sets `false` to restore the visible default.

## Risks / Trade-offs

- [A forged capability attempts to edit an opaque root] -> Require a capability
  whose expected hash matches the fresh page object, then let the codec re-prove
  visibility semantic hash and lexical profile.
- [Page hidden is confused with element hidden] -> Keep one operation name but
  issue it on different nativeRef scopes; the field path remains `hidden`.

## Migration Plan

Additive schema vocabulary only. Existing PPJ remains valid.

## Open Questions

None.
