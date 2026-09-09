## Context

See proposal.md. BuildTextStyle supplies FontFamilyEastAsia from FontFamily when the former is missing. Source edits currently rebuild that authored model, so copying its inferred presence would defeat explicit deletion.

## Goals / Non-Goals

Complete a:ea presence for ordinary text/shape paragraphs. Keep authored fallback, other script fonts, placeholder inheritance and host substitution separate.

## Decisions

Extend exact field masking/authority and clone the source default style. For East Asian changes, clear the field and copy the resolved value only when the raw requested paragraph explicitly contains fontFamilyEastAsia. Reuse existing token resolution and nonblank/255-character validation.

Patch only the changed East Asian font through the existing guarded writer, and inspect serialized leaf content before typed access. Reuse the shared font lifecycle/rejection fixtures for a:latin and a:ea, including Latin-retained deletion that would expose fallback reintroduction.

## Risks / Trade-offs

Authored fallback can mask deletion → require raw presence and assert absence after fresh projection. Unknown native font metadata/children can be lost → reject replacement and check source bytes. Preserve the documented independent whole-default-style baseline failure exclusion.
