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
    desiredOutcome: "Approve the bounded migration",
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
    designGrammar: {
      palette: {
        roles: { background: "#F5F1E8", ink: "#17202A", accent: "#D15B2A" },
        strict: false,
      },
      typography: {
        roles: { title: "Aptos Display", body: "Aptos", data: "Aptos Mono" },
        minimumBodyFontSize: 20,
        minimumCaptionFontSize: 18,
        strict: false,
      },
      spacing: { base: 8, outerMargin: 48 },
      densityRhythm: "alternate sparse decisions with denser evidence",
      motif: "thin route lines and one cropped evidence frame",
      imagery: "documentary, task-relevant, minimal decoration",
      charts: "direct labels, shared scale, one highlighted series",
      invariants: ["one primary claim per page", "source note on evidence pages"],
      intentionalWarnings: [{
        type: "cardWallPattern",
        pageIds: ["page-02-options"],
        reason: "Equal-sized options are deliberately compared on one shared decision axis.",
      }],
    },
  },
  pages: [{
    id: "page-01-decision",
    readerTask: "Understand the decision",
    claim: "The bounded path is the safest deployable option",
    evidence: ["Recovery test result"],
    contentBudget: { maxCharacters: 420, maxObjects: 12 },
    compositionIntent: "one decision statement above a compact option comparison",
    assetRefs: [],
  }],
  editorial: {
    voice: "direct, specific, evidence-led",
    titleMode: "claim",
    lockedFacts: ["Measured values and source labels"],
    avoid: ["empty transitions", "stacked slogans", "false certainty"],
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

## Artifact references

Declare a managed artifact once:

```js
artifactRefs: [{
  artifactId: "brand-template",
  sha256: "<64 lowercase hex characters>",
  role: "authoritative-template",
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
