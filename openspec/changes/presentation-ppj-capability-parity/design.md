## Context

PPJ 2.0 established one declarative public Presentation language, but its first
release intentionally concentrated on the complete compiler path rather than
exhaustive capability ergonomics. Since the Presentation import branch forked,
the codec has gained 93 commits covering bounded fill, line, text-body,
paragraph, bullet, and run-style leaves. The current PPJ schema is already
large, while its generated reference collapses most nested definitions into a
short top-level table and its method-oriented registry does not describe schema
paths or native leaf kinds.

The integration must preserve two distinct truths:

- source-free PPJ fields are authored semantic state and must compile natively;
- imported native leaves are revision-bound capabilities and must token-splice
  only the field issued by the exact source projection.

## Goals / Non-Goals

**Goals:**

- integrate the stable import-primitive branch without restoring a public JS
  Presentation authoring route;
- make every accepted PPJ field and issued native leaf discoverable;
- ensure an accepted authored field is compiled rather than silently ignored;
- let an Agent edit bounded imported native leaves in PPJ without raw package
  locators;
- make future primitive additions fail maintenance checks when PPJ ownership is
  missing.

**Non-Goals:**

- model arbitrary OOXML or make every preserved native graph editable;
- expose raw XML, XPath, relationship IDs, part names, or arbitrary properties;
- copy every legacy facade method into PPJ as a command;
- restore MJS/Compose as a public authoring language;
- add a new wire version, authoring engine, or exhaustive visual benchmark.

## Decisions

### Keep PPJ state-oriented

New authored support is represented as typed state under existing element,
text, paragraph, fill, stroke, media, chart, table, motion, and deck structures.
Legacy methods remain classified as helpers, inspection, or host operations.
Serializing method calls was rejected because it would recreate the imperative
state and attention burden PPJ was introduced to remove.

### Add a closed native-leaf representation

`nativeRef` gains an optional ordered `leaves` collection. Each leaf carries a
stable opaque leaf ID, a closed `kind`, a typed scalar `value`, and an expected
hash. The PPJ contains no package locator. On build, C# reprojects the exact
source, verifies the source/object/capability/leaf hashes, compares values, and
lowers only changed leaves through the existing edit-plan codec.

Kinds are additive but closed in the schema and C# validator. Value domains are
declared in the capability registry and revalidated by the native codec. A
generic property bag was rejected because it would turn PPJ into a raw attribute
patch language.

### Make the registry field-level

`src/ppj/capability-registry.json` remains the maintainer-owned index and gains:

- `ppjPaths`: schema paths with authored/projected/review ownership;
- `nativeLeafKinds`: issued kind, scalar type, value domain, PPJ location,
  lowering owner, and reference section;
- explicit `internal` and `hostOnly` classifications.

The JSON Schema remains the syntax authority; the registry records operational
ownership rather than duplicating the schema. A generator resolves the schema
and registry together to produce the Agent reference.

### Generate a complete progressive reference

The generated `ppj.md` keeps a concise workflow and element index first, then
includes exhaustive nested type/property/value tables, authored examples, and a
native-leaf capability table. Agents can read the beginning for routing and
search the same file for a precise field or leaf. Maintaining a short manual and
a separate exhaustive manual was rejected because they would drift.

### Integrate history, preserve PPJ authority

The stable primitive branch is merged with its atomic history. Conflicts in
public Help, Skill routing, release metadata, and package policy are resolved in
favor of PPJ 2.0; codec, protobuf, source-bound proof, and real six-sample
evidence are retained. Work that was not pushed at the frozen branch SHA is not
included.

### Validate behavior, not every parameter

One comprehensive authored text/line program and one imported multi-leaf edit
protect the public contract. Existing six-sample evidence remains the native
oracle. No effect matrix, per-enum test duplication, benchmark harness, or full
release run is added during development.

## Risks / Trade-offs

- **[Risk] Schema says more than the authored compiler implements.** → Generate
  the path inventory and make the parity check require an authored compiler
  owner or an explicit metadata-only classification.
- **[Risk] A generic native leaf becomes an OOXML escape hatch.** → Keep leaf
  kinds closed, opaque IDs hash-bound, scalar-only, and re-proven by the native
  edit-plan codec.
- **[Risk] Merging the long-lived branch revives retired JS documentation.** →
  resolve routing/documentation conflicts in favor of PPJ and scan the public
  Skill for retired authoring instructions.
- **[Risk] A complete reference consumes Agent context.** → Put routing and
  common examples first; keep exhaustive tables searchable and generated.
- **[Trade-off] Some codec capabilities remain nativeRef-only.** → Prefer honest
  bounded editability over pretending that a partial source token is a complete
  authored semantic model.

## Migration Plan

1. Freeze and integrate the pushed primitive branch SHA.
2. Regenerate protobuf bindings and make the native build green.
3. Extend schema/model/projector/lowerer and the field-level registry.
4. Regenerate `ppj.md`, API evidence, and capability checks.
5. Run narrow PPJ and Presentation checks, commit atomically, then fast-forward
   the integration branch into `main`.

Rollback is a normal revert of the additive integration commits. Existing PPJ
files remain valid because all language changes are optional and additive.

## Open Questions

None. Newly discovered partial semantics default to `nativeRef` until an
authored typed contract is independently complete.
