## Context

Kimi's local presentation package is intentionally layered. Its main
`kimi-slides` Skill chooses the task and references; `presentation-artifact-tool`
owns the PPTD/runtime workflow; `pptd.md`, `shapes.md`, `fonts.md`, and the
scenario documents are contracts or guidance loaded on demand. The image
search and browser Skills are optional integrations, not part of the PPTD
language. This is useful architecture, but Kimi's `.pptd` and private services
are not OfficeKit dependencies or a reason to copy its text.

OfficeKit already has a stronger native source-bound model and many more
bounded operations. Its weakness is discoverability: Help, generated API docs,
the Presentations router, task routes, template Creator, and examples do not
share an explicit ownership map. Some long references repeat the same
workflow, while the primitive surface is not presented as one learnable
language.

## Goals / Non-Goals

**Goals:**

- Make the distinction between format/runtime facts, semantic primitives,
  authoring guidance, task routing, and optional providers explicit.
- Give an Agent one short Presentation entry point and one canonical semantic
  primitive index that points to authoritative API/Help details.
- Give typography and shape guidance the same first-class, role-based status as
  scenarios and motion.
- Provide a lightweight `skill-update` Skill/checker that tells maintainers
  which surfaces must change after a runtime primitive change.
- Preserve native OfficeKit capabilities, source-bound safety, and clean-room
  boundaries while removing duplicated instructions.

**Non-Goals:**

- No `.pptd` language, universal AST, Office wire change, codec rewrite, or new
  public Presentation API.
- No copy of Kimi source text, private protocol, remote template payload, or
  proprietary asset.
- No automatic prose or code editing by the update checker, no benchmark
  matrix, and no broad test-suite expansion.
- No deletion of a detailed reference until its information has one clear
  owner and a route points to that owner.

## Decisions

### 1. Use five information layers

The Presentation package will use this authority order:

1. **Runtime/API facts** — public JS exports, `src/help/index.mjs`, and
   generated `docs/api.md` are authoritative for signatures and observable
   boundaries.
2. **Primitive language** — one concise `references/primitives.md` groups the
   API into semantic families and links to the runtime facts; it never restates
   every option.
3. **Task route** — `SKILL.md` and `tasks/*.md` tell an Agent what to load and
   what sequence to follow.
4. **Design guidance** — doctrine, scenario, typography, imagery, motion, and
   review references explain choices and anti-patterns, not API facts.
5. **Provider/integration guidance** — image sourcing, live hosts, and PDF or
   spreadsheet owners remain separate Skills and are loaded only when needed.

This keeps Kimi's useful separation without forcing its intermediate format
onto OfficeKit.

### 2. Make the impact manifest the maintenance contract

`skills/presentations/skills/presentations/references/primitive-impact.json`
will map stable primitive families to source globs, Help/API output, the
semantic reference, route/reference files, examples, focused tests, and
coverage/release evidence. A family can list an explicit `notes` boundary when
an operation is intentionally source-bound or fail-closed. The manifest is
small and reviewed with primitive changes; it is not generated from private
runtime internals.

### 3. Keep `skill-update` advisory and deterministic

The new Skill exposes a script with `impact` and `check` commands. `impact`
matches changed paths against the manifest and prints the affected surfaces;
`check` verifies that every manifest path exists, every referenced public Help
name is present, every route/reference link resolves, and no Kimi-only path or
identifier is packaged. It reads git state and text files only. It never
downloads providers, initializes Office WASM/NativeAOT, edits files, or claims
that a primitive is implemented merely because documentation exists.

### 4. Refactor by ownership, not by mass deletion

The first implementation makes the router and primitive/typography references
authoritative and turns known duplicate references into short redirects. The
long imported-editing contract remains available for advanced source-bound
work, but its repeated creation, template, and delivery prose is removed only
where the route now owns it. This avoids breaking existing tasks while making a
new Agent see a coherent path.

### 5. Share guidance with the Template Creator through links

The Creator Skill loads the primitive and typography references when it builds
calibration examples, and records capability evidence in its existing package
spec. It does not copy the Presentation route or promise source-graph
editability in a style template.

## Risks / Trade-offs

- [The primitive index becomes a second API specification] → Keep signatures
  out of it, link each family to Help/API, and make the checker reject unknown
  Help names rather than duplicating option tables.
- [Aggressive shortening hides an advanced safety boundary] → Preserve
  source-bound/fail-closed notes in the index and leave advanced details in the
  referenced contract.
- [A static impact map becomes stale] → `skill-update check` verifies paths and
  Help names; primitive commits must update the map in the same atomic change.
- [Kimi-inspired language leaks proprietary material] → Keep only the
  clean-room classification and write OfficeKit-specific guidance from public
  behavior and standards.
- [Installing another Skill increases clutter] → Keep `skill-update` one small
  host-neutral Skill with no extra runtime dependency or task files.

## Migration Plan

1. Add the OpenSpec contract, impact manifest, and `skill-update` Skill.
2. Add the primitive and typography references; shorten the Presentation and
   Creator routers to link them.
3. Update Help metadata only where a missing adoption link is a real discovery
   gap, then regenerate API docs if the Help catalog changes.
4. Run the checker and existing narrow Skill smoke; use one fresh-context
   presentation task to confirm the route finds primitives without loading the
   full API corpus.
5. Commit the reform by subsystem and merge onto the latest `main` only after
   the worktree contains no user WIP.

Rollback is a normal revert of the new Skill/docs commits. Runtime and wire
behavior remain unchanged.

## Open Questions

None for this slice. Whether a future `DeckPlan` compiler or a richer native
primitive DSL is useful remains a later product decision, not an implicit part
of this reform.
