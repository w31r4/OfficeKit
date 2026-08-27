# Presentation authoring plan

Use one durable plan to carry communication intent across REPL sessions. The
plan is orchestration state; MJS/Compose remains executable authoring source.

## Contract

- Schema: `office-kit/presentation-authoring-plan/v1`
- Maximum canonical JSON: 256 KiB
- Maximum pages: 64
- Storage: private immutable `plans/<sha256>.json`
- Update: pass the exact active `expectedSha256`
- References: use task artifact ID and SHA-256

The plan contains plain JSON only. Keep functions, runtime objects, model
transcripts, raw OOXML, XPath, part paths, relationship IDs, and source bytes
outside it.

## Required shape

```js
const plan = {
  schema: "office-kit/presentation-authoring-plan/v1",
  mode: "create",
  brief: {
    audience: "Engineering leadership",
    purpose: "Choose a deployment architecture",
    primaryJob: "decide",
    supportingJobs: ["align"],
    expectedOutcome: "Leadership approves the bounded migration and its owner",
    mediumFit: "strong",
    afterUse: "Decision record and implementation handoff",
    durationMinutes: 15,
    language: "en",
    evidence: ["Measured build and recovery results"],
    constraints: ["Keep all figures traceable"],
  },
  narrative: {
    thesis: "The bounded path reduces operational risk without delaying delivery",
    sections: ["Current constraint", "Evidence", "Decision", "Plan"],
  },
  design: {
    sourceMode: "self-directed",
    mechanismPacks: ["technical-architecture"],
    scenario: {
      primary: "technical-engineering",
      secondary: "analysis-decision",
    },
    direction: {
      name: "Traceable migration decision",
      rationale: "The audience needs a system view tied directly to evidence, risk, and ownership",
    },
    designGrammar: {
      palette: {
        roles: { background: "#F5F1E8", ink: "#17202A", accent: "#D15B2A" },
        strict: false,
      },
      typography: {
        roles: { title: "Aptos Display", body: "Aptos", data: "Aptos Mono" },
        minimumBodyFontSize: 22,
        minimumCaptionFontSize: 20,
        strict: false,
      },
      spacing: { base: 8, outerMargin: 48 },
      geometry: "straight boundaries, square evidence frames, restrained rounding",
      lines: "solid for control flow, dashed for deferred migration",
      densityRhythm: "alternate sparse decisions with denser evidence",
      visualCarriers: ["system boundary diagram", "shared-scale comparison chart"],
      motif: "thin route lines and one cropped evidence frame",
      imagery: "documentary, task-relevant, minimal decoration",
      charts: "direct labels, shared scale, one highlighted series",
      antiPatterns: ["generic card wall", "unexplained arrows", "decorative circles"],
      invariants: ["one primary claim per page", "source note on evidence pages"],
      intentionalWarnings: [{
        type: "cardWallPattern",
        pageIds: ["page-02-options"],
        reason: "Equal-sized options are deliberately compared on one shared decision axis.",
      }],
    },
    calibration: {
      pageIds: ["page-01-decision", "page-03-evidence", "page-05-risk"],
      status: "reviewed",
      decision: "Keep the shared scale and increase evidence-page label contrast.",
    },
  },
  pages: [{
    id: "page-01-decision",
    readerTask: "Understand the decision",
    claim: "The bounded path is the safest deployable option",
    evidence: ["Recovery test result"],
    contentBudget: { maxCharacters: 420, maxObjects: 12 },
    compositionIntent: "shared-scale comparison chart from measured task data, with one decision statement above it",
    assetRefs: [],
  }],
  editorial: {
    voice: "direct, specific, evidence-led",
    titleMode: "claim",
    lockedFacts: ["Measured values, dates, units, and source labels"],
    protectedWording: ["Approved product and legal terminology"],
    avoidPatterns: ["throat clearing", "stacked slogans", "false certainty"],
    scope: { mode: "deck", pageIds: [] },
  },
  artifactRefs: [],
  recipe: "tasks/create.md",
  unresolved: [],
  nextAction: "Compose and review the first complete working draft",
};
```

For a self-directed deck, keep `palette.strict: false`. Reserve `strict: true`
for an authoritative closed brand palette, and enumerate every emitted color,
including theme/default colors, before binding the plan.

New plans require `brief.primaryJob`, `brief.expectedOutcome`, one primary
`design.scenario`, and a selected `design.direction`. When `mediumFit` is
`weak`, include `mediumFitNote` with the limitation and mitigation. Existing
pre-strategy plans remain readable and appear as `strategyStatus: "legacy"` in
task summaries; do not manufacture missing strategy during an unrelated local
edit.

The `editorial` object uses the existing open JSON contract; it does not create
a second schema. Record the intended `voice`, `titleMode`, `lockedFacts`,
`protectedWording`, `avoidPatterns`, and `scope`. Use `scope.mode: "local"`
with exact plan page IDs for a bounded edit and `scope.mode: "deck"` only when
the task authorizes deck-wide shaping. The sibling
`presentation-editorial-trim` Skill interprets these fields while preserving
names, values, citations, uncertainty, and source meaning.

For self-directed decks longer than four pages, record the opening, evidence,
and densest/high-risk calibration page IDs under `design.calibration`. A deck
of four pages or fewer records the complete page set. When calibration changes
the grammar, write an updated plan with the current `expectedSha256`; resume
uses that plan revision and the latest reviewed artifact rather than a second
design file or restored JavaScript heap.

Each page `compositionIntent` must name the visual carrier and its source
strategy. It should distinguish, for example, a native chart from measured task
data, a supplied image, a style-guided free composition, a source-derived
component from a reference deck, or an authored diagram.
Coordinates, columns, and “clean layout” are implementation details, not
sufficient intent.

## Artifact references

Declare a managed artifact once:

```js
artifactRefs: [{
  artifactId: "brand-reference-deck",
  sha256: "<64 lowercase hex characters>",
  role: "authoritative-reference-deck",
}]
```

Reference it from another field with the same pair under `artifactRef`.
OfficeKit verifies that the task owns that revision.

## Write, read, and update

```js
const first = await ctx.plan(plan);
const current = await ctx.plan();

const revised = structuredClone(current);
revised.pages[0].claim = "A sharper, evidence-backed decision";
const next = await ctx.plan(revised, { expectedSha256: first.sha256 });
```

An identical canonical write is idempotent. A missing or stale hash leaves the
active plan unchanged. After a changed write, create and review a new artifact
commit before publication.

## Unresolved decisions

Use a string or an object. Strings and entries whose `required`/`blocking`
flags are not false block design review. Mark intentionally deferred work as:

```js
{ id: "optional-photo", required: false, note: "Use native shapes if unavailable" }
```
