---
name: presentation-skill-maintainer
description: Keep OfficeKit PPJ, native compiler capabilities, Presentation Help, Agent guidance, review rules, and examples synchronized. Use when changing a PPJ field, Presentation codec/runtime capability, imported nativeRef operation, review invariant, or presentation Skill route.
---

# Presentation Skill Maintainer

Use this Skill for repository maintenance, not for authoring a user deck. PPJ
has one language contract and the Presentations Skill has one Agent route; a
runtime change is incomplete when either surface cannot discover it.

## Start with the owning class

Read `src/ppj/capability-registry.json`. It is the discoverability ledger for
three surfaces: PPJ root paths, the closed C# source-edit leaf vocabulary, and
every stable Presentation Help API. Help APIs are classified as one of:

- `ppj-state`: persistent state belongs in the JSON Schema and compiler;
- `native-ref`: a source-bound capability belongs in projection, diff and
  fail-closed lowering;
- `compiler-helper`: an implementation convenience, not Agent syntax;
- `inspect-review`: discovery or evidence, not persistent page state;
- `host-only`: an open-PowerPoint action that never enters PPJ.

Do not expose a sixth escape hatch. Raw OOXML, XPath, relationship IDs and
arbitrary JavaScript are not Presentation language features.

## Propagate one change

1. Update the authoritative runtime or protocol implementation.
2. Update `src/ppj/ppj-v1.schema.json` when persistent state changes.
3. Update the C# reader, validator, projector, compiler and lowerer surfaces
   that own the classified capability.
4. Update Help and the capability registry. Classify every new stable Help API,
   PPJ owner path and `PptxEditPlanCodec` leaf explicitly; the checker rejects
   an orphan in either direction.
5. Run `sync` to regenerate `references/ppj.md` from the Schema and registry.
6. Update exactly one focused Agent reference: creative direction, fonts,
   shapes, text, charts/tables, media/layers, motion, components/templates,
   imported nativeRef, scenarios, or review/delivery.
7. Add one minimal example or regression only for a stable contract, material
   risk or reproduced failure. Do not create a parameter matrix for coverage.
8. Record unverified renderer or host behavior honestly.

## Commands

From the repository root:

```bash
node skills/presentations/skills/presentation-skill-maintainer/scripts/maintain-presentation-skill.mjs check
node skills/presentations/skills/presentation-skill-maintainer/scripts/maintain-presentation-skill.mjs sync
```

`check` is read-only. It compares the registry with every Presentation Help
record and the codec's closed leaf set, verifies every PPJ root path has an
owner, keeps PowerPoint Live host-only, and proves the exhaustive generated PPJ
manual matches its inputs. `sync` changes only the generated PPJ manual.

## Completion boundary

A green checker proves discoverability and documentation synchronization. It
does not prove codec behavior, visual quality, host playback, licensing or
release readiness. Run the narrow owning check separately and keep those
evidence claims distinct.
