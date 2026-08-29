#!/usr/bin/env node

import { createHash } from "node:crypto";
import { existsSync, readFileSync, writeFileSync } from "node:fs";
import path from "node:path";
import process from "node:process";

const repo = path.resolve(process.cwd());
const command = process.argv[2] ?? "check";
if (!new Set(["check", "sync"]).has(command)) throw new Error(`Unknown command: ${command}`);

const registryPath = path.join(repo, "src/ppj/capability-registry.json");
const schemaPath = path.join(repo, "src/ppj/ppj-v1.schema.json");
const registry = readJson(registryPath);
const schema = readJson(schemaPath);
const target = path.join(repo, registry.generatedReference);
const errors = validateRegistry(registry, schema);
const generated = renderManual(schema, registry);

if (command === "sync") {
  if (errors.length) fail(errors);
  writeFileSync(target, generated);
  process.stdout.write(`Synchronized ${path.relative(repo, target)}\n`);
} else {
  if (!existsSync(target)) errors.push(`Missing generated PPJ reference: ${path.relative(repo, target)}`);
  else if (readFileSync(target, "utf8") !== generated) errors.push("Generated ppj.md is stale; run the maintainer sync command.");
  if (errors.length) fail(errors);
  process.stdout.write(`Presentation Skill maintenance check ok · ${Object.keys(registry.helpApis).length} Help APIs · ${registry.hostOnly.length} host-only operations\n`);
}

function validateRegistry(value, language) {
  const errors = [];
  if (value.schema !== "office-kit/presentation-capability-registry/v1") errors.push("Unexpected capability registry schema.");
  if (language.properties?.schema?.const !== "office-kit/ppj/v1") errors.push("Unexpected PPJ language schema.");
  const classes = new Set(["ppj-state", "native-ref", "compiler-helper", "inspect-review", "host-only"]);
  if (JSON.stringify([...Object.keys(value.classes ?? {})].sort()) !== JSON.stringify([...classes].sort())) errors.push("Capability classes are incomplete.");
  const helpSource = readFileSync(path.join(repo, value.helpSource), "utf8");
  const helpNames = [...helpSource.matchAll(/\{\s*artifactKind:\s*"presentation"\s*,\s*kind:\s*"api"\s*,\s*name:\s*"([^"]+)"/g)]
    .map((match) => match[1]).sort();
  const registryNames = Object.keys(value.helpApis ?? {}).sort();
  for (const name of helpNames) if (!(name in value.helpApis)) errors.push(`Orphan Presentation Help API: ${name}`);
  for (const name of registryNames) if (!helpNames.includes(name)) errors.push(`Registry API has no current Help record: ${name}`);
  for (const [name, classification] of Object.entries(value.helpApis ?? {})) if (!classes.has(classification)) errors.push(`Invalid class for ${name}: ${classification}`);

  const liveSource = readFileSync(path.join(repo, "src/live/adapters/powerpoint.mjs"), "utf8");
  const liveBlock = /POWERPOINT_LIVE_OPERATIONS\s*=\s*Object\.freeze\(\[([\s\S]*?)\]\)/u.exec(liveSource)?.[1] ?? "";
  const liveNames = [...liveBlock.matchAll(/"([a-z_]+)"/g)].map((match) => `powerpoint.${match[1]}`).sort();
  const declaredLive = [...(value.hostOnly ?? [])].sort();
  if (JSON.stringify(liveNames) !== JSON.stringify(declaredLive)) errors.push("PowerPoint Live operations and host-only registry entries differ.");
  for (const name of value.hostOnly ?? []) if (value.helpApis?.[name] != null) errors.push(`Host-only operation leaked into file authoring Help: ${name}`);
  return errors;
}

function renderManual(schema, registry) {
  const schemaBytes = readFileSync(schemaPath);
  const registryBytes = readFileSync(registryPath);
  const elementRefs = schema.$defs.element.oneOf.map((entry) => entry.$ref.split("/").at(-1));
  const rootDescriptions = {
    schema: "Fixed language identifier `office-kit/ppj/v1`.",
    meta: "Stable program identity, title, language and revision version.",
    intent: "Audience, brief, narrative, editorial constraints and delivery purpose.",
    design: "Canvas, theme, named styles, deck-specific Design Grammar and motion policy.",
    assets: "Content-addressed local resources with rights and accessibility metadata.",
    source: "Optional immutable third-party PPTX binding for source-preserving edits.",
    components: "Finite reusable structures with parameters, slots, variants, repeat and simple conditions.",
    pages: "Ordered slides; each page's element array is the real back-to-front z-order.",
    sections: "Persistent section membership over stable page IDs.",
    customShows: "Named ordered page subsets.",
    comments: "Persistent presentation comments supported by the language contract.",
  };
  const required = new Set(schema.required ?? []);
  const rootRows = Object.keys(schema.properties).map((name) =>
    `| \`${name}\` | ${required.has(name) ? "yes" : "no"} | ${rootDescriptions[name]} |`).join("\n");
  const elementRows = elementRefs.map((name) => {
    const definition = schema.$defs[name];
    const fragments = (definition.allOf ?? [definition]).map((fragment) => fragment.$ref ? schema.$defs[fragment.$ref.split("/").at(-1)] : fragment);
    const properties = Object.assign({}, ...fragments.map((fragment) => fragment.properties ?? {}));
    const requiredFields = new Set(fragments.flatMap((fragment) => fragment.required ?? []));
    const type = properties.type?.const ?? name.replace(/Element$/u, "");
    const mandatory = [...requiredFields].filter((field) => field !== "type").map(code).join(", ") || "none";
    const optional = Object.keys(properties).filter((field) => !requiredFields.has(field) && field !== "type").map(code).join(", ") || "none";
    return `| \`${type}\` | ${mandatory} | ${optional} |`;
  }).join("\n");
  const counts = Object.entries(registry.helpApis).reduce((value, [, classification]) => {
    value[classification] = (value[classification] ?? 0) + 1;
    return value;
  }, {});
  const classRows = Object.entries(registry.classes).map(([name, details]) =>
    `| \`${name}\` | ${counts[name] ?? registry.hostOnly.length} | ${details.meaning} |`).join("\n");
  const budgets = schema["x-officekit-budgets"];
  return `<!-- GENERATED by presentation-skill-maintainer; do not hand-edit. schema-sha256=${digest(schemaBytes)} registry-sha256=${digest(registryBytes)} -->
# PPJ language reference

PPJ is OfficeKit's single public presentation authoring language. It is one
UTF-8 strict JSON file with schema \`office-kit/ppj/v1\`. Edit the file, then
use the CLI to validate, build, render and review it. JavaScript may generate
JSON externally, but JavaScript functions, JSON5, raw OOXML, XPath, relationship
IDs, network calls and executable expressions are not PPJ.

## Workflows

\`\`\`text
new deck:       deck.ppj → check → build → render → review
third-party:    input.pptx → import → inspect/edit deck.ppj → build
OfficeKit PPTX: authored.pptx → import → exact embedded PPJ recovery
durable task:   add --task only when immutable revision/resume evidence is wanted
\`\`\`

The CLI is \`officekit ppj import|inspect|check|build|render|review\`. Build
never overwrites the PPJ or its bound source. Render and review are explicit;
successful compilation is not visual approval.

## Root fields

| Field | Required | Meaning |
| --- | --- | --- |
${rootRows}

All objects are closed: undeclared fields fail validation. IDs match
\`${schema.$defs.id.pattern}\` and remain stable across edits. Coordinates and
sizes use points. Page order is \`pages[]\` order. Element order is back-to-front
\`pages[].elements[]\` order; do not invent a second z-index.

## Typed page elements

| \`type\` | Required fields beyond \`type\` | Optional fields |
| --- | --- | --- |
${elementRows}

Simple text uses a string. Mixed formatting uses \`paragraphs[]\` and
\`runs[]\`; do not encode markup inside strings. Colors use theme references or
explicit typed color objects. Assets use relative URIs, exact MIME and SHA-256;
remote URLs and data-fetch instructions are invalid. Accessibility and rights
metadata travel with the asset or element.

## Components terminate

Components have a finite frame, typed parameters, named slots, explicit
variants, bounded repeat items and only \`equals\`, \`notEquals\`, \`present\`
or \`absent\` conditions. They cannot call themselves recursively. Expanded
IDs are deterministic. A component is reuse, not a hidden script.

## Imported PPTX and nativeRef

Import copies the source package into a content-addressed local asset and binds
its SHA-256. Every visible object becomes a typed element or \`opaque\` with a
\`nativeRef\`. A nativeRef lists only capability-issued fields. Edit those
fields and keep its expected revision/hash; unsupported topology stays opaque.
No-op build returns the source bytes exactly. A stale, ambiguous or undeclared
mutation fails instead of rebuilding or flattening the source.

OfficeKit-authored PPTX embeds canonical PPJ and a node map. Import restores
that PPJ exactly when valid. If native software changed the PPTX but left the
embedded program, PPJ remains authoritative; a future build writes a new output
and never overwrites the input.

## Hard budgets

| Budget | Limit |
| --- | ---: |
| PPJ UTF-8 bytes | ${budgets.maxSourceBytes} |
| pages | ${budgets.maxPages} |
| expanded elements | ${budgets.maxExpandedElements} |
| one repeat | ${budgets.maxRepeatItems} |
| component expansion depth | ${budgets.maxComponentDepth} |

Budget, reference, type, cycle and source-capability errors are reported before
compilation. \`ppj check --fix\` may normalize deterministic formatting; it
must not choose layout, rewrite copy or change design semantics.

## Minimum authored program

\`\`\`json
{
  "schema": "office-kit/ppj/v1",
  "meta": { "id": "brief", "title": "Decision brief", "language": "en-US", "version": 1 },
  "intent": {},
  "design": {},
  "pages": [
    {
      "id": "opening",
      "elements": [
        { "type": "text", "id": "claim", "frame": { "x": 48, "y": 42, "width": 624, "height": 72 }, "text": "Evidence changed the decision" }
      ]
    }
  ]
}
\`\`\`

## Capability ownership

| Class | Current entries | Meaning |
| --- | ---: | --- |
${classRows}

The registry classifies legacy facade APIs while PPJ 2.0 converges. A
\`compiler-helper\` is not Agent syntax. PowerPoint Live operations remain in
the separate host-only list and never serialize into PPJ.

## Common mistakes

- Editing a PPTX package path instead of its PPJ ID or issued nativeRef.
- Reordering a type-specific collection instead of the page element array.
- Putting base64, HTTP URLs, functions or expressions in the program.
- Treating opaque-preserved as editable, or rebuilding the whole deck after a
  rejected source-bound edit.
- Calling build success a render, visual review or PowerPoint playback result.
- Using components as an unbounded layout engine instead of explicit finite
  reuse.
`;
}

function code(value) { return `\`${value}\``; }
function digest(value) { return createHash("sha256").update(value).digest("hex"); }
function readJson(file) { return JSON.parse(readFileSync(file, "utf8")); }
function fail(errors) { for (const error of errors) process.stderr.write(`ERROR ${error}\n`); process.exit(1); }
