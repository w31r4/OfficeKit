## Context

The shared PPJ shadow model and authored compiler already carry `rotateWithShape`, while the imported direct-run profile currently rejects that attribute. Source-bound edits for neighboring shadow fields already prove the owning `a:outerShdw` and splice a single XML token. The change follows those boundaries and keeps unsupported effect topology opaque.

## Goals / Non-Goals

**Goals:**

- Admit `rotWithShape` as one allowed direct-shadow transform and expose it as a boolean native leaf.
- Preserve strict direct-run topology checks and a token-only source-bound edit.
- Keep the PPJ wire shape, host adapters, and authored shadow construction unchanged.

**Non-Goals:**

- Supporting combinations of rotation with scale or skew transforms.
- Rewriting complex effect lists, WordArt/default text, shape shadows, or host rendering.
- Changing the Office wire version or claiming PowerPoint behavioral acceptance.

## Decisions

1. **Count rotation as a single allowed transform.** Extend the existing direct-shadow safety predicates with an opt-in `allowRotateWithShape` flag and include rotation in the single-transform count. This preserves current rejection defaults and avoids accepting scale/skew combinations. A separate parser would duplicate the established geometry/color/unknown-child checks.

2. **Use a boolean native leaf and token splice.** The leaf maps to `run.style.shadow.rotateWithShape`; the edit dispatcher maps it to `@rotWithShape` and reuses the existing direct-shadow owner proof. This keeps explicit `0` and `1` round-trippable without reconstructing the effect XML.

3. **Regenerate the public capability surfaces.** Update the schema, registry, generated matrix/reference, coverage, backlog, and presentation Skill together so the typed field and its bounded evidence have one discoverable contract. No protobuf change is needed because the shared shadow field already exists.

## Risks / Trade-offs

- [Risk] OpenXML accepts several boolean spellings while the source-bound patch requires a stable token. → Mitigation: the bounded profile admits only canonical `0`/`1`; other spellings remain opaque or fail closed.
- [Risk] A future effect combination may be mistaken for a safe direct shadow. → Mitigation: rotation is counted with scale/skew and the existing sibling/unknown/geometry/color checks remain mandatory.
- [Risk] Focused codec tests do not prove host PowerPoint rendering. → Mitigation: document the experiment as source/projection evidence only and leave host acceptance outside this increment.

## Migration Plan

No data migration is required. Existing PPJ documents remain valid; imported runs gain the leaf only when they satisfy the new bounded profile. The change can be rolled back as one commit without changing the wire version.
