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
const presetProfilesPath = path.join(repo, registry.presetGeometryProfiles);
const presetProfiles = readJson(presetProfilesPath);
const target = path.join(repo, registry.generatedReference);
const errors = validateRegistry(registry, schema, presetProfiles);
const generated = renderManual(schema, registry, presetProfiles);

if (command === "sync") {
  if (errors.length) fail(errors);
  writeFileSync(target, generated);
  process.stdout.write(`Synchronized ${path.relative(repo, target)}\n`);
} else {
  if (!existsSync(target)) errors.push(`Missing generated PPJ reference: ${path.relative(repo, target)}`);
  else if (readFileSync(target, "utf8") !== generated) errors.push("Generated ppj.md is stale; run the maintainer sync command.");
  if (errors.length) fail(errors);
  process.stdout.write(
    `Presentation Skill maintenance check ok · ${Object.keys(registry.helpApis).length} Help APIs · ` +
    `${Object.keys(registry.nativeLeafKinds).length} native leaves · ${registry.hostOnly.length} host-only operations\n`,
  );
}

function validateRegistry(value, language, geometryProfiles) {
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

  const expectedPaths = Object.keys(language.properties ?? {}).map((name) => `$.${name}`).sort();
  const declaredPaths = Object.keys(value.ppjPathOwners ?? {}).sort();
  if (JSON.stringify(expectedPaths) !== JSON.stringify(declaredPaths))
    errors.push("PPJ root fields and ppjPathOwners entries differ.");
  for (const [field, details] of Object.entries(value.ppjPathOwners ?? {})) {
    if (!details || typeof details.owner !== "string" || typeof details.surface !== "string" || typeof details.meaning !== "string")
      errors.push(`PPJ field owner ${field} is incomplete.`);
  }
  const ppjStateApis = Object.entries(value.helpApis ?? {})
    .filter(([, classification]) => classification === "ppj-state")
    .map(([name]) => name)
    .sort();
  const mappedStateApis = Object.keys(value.ppjStateApiPaths ?? {}).sort();
  if (JSON.stringify(ppjStateApis) !== JSON.stringify(mappedStateApis))
    errors.push("PPJ-state Help APIs and ppjStateApiPaths entries differ.");
  for (const [name, ppjPath] of Object.entries(value.ppjStateApiPaths ?? {})) {
    if (typeof ppjPath !== "string" || !ppjPath.startsWith("$.")) {
      errors.push(`PPJ-state API ${name} has an invalid PPJ path.`);
      continue;
    }
    const rootName = /^\$\.([A-Za-z][A-Za-z0-9]*)/u.exec(ppjPath)?.[1];
    if (!rootName || !(rootName in (language.properties ?? {})))
      errors.push(`PPJ-state API ${name} references unknown root path ${ppjPath}.`);
  }
  if (geometryProfiles.schema !== "office-kit/ppj-preset-geometry-profiles/v1")
    errors.push("Unexpected preset geometry profile schema.");
  const geometrySchema = language.$defs.geometry.oneOf.find((entry) => entry.properties?.kind?.const === "preset");
  const presetNames = [...(geometrySchema?.properties?.preset?.enum ?? [])].sort();
  const profileNames = Object.keys(geometryProfiles.profiles ?? {}).sort();
  if (JSON.stringify(presetNames) !== JSON.stringify(profileNames))
    errors.push("PPJ preset geometry enum and canonical profile registry differ.");
  const adjustmentItems = geometrySchema?.properties?.adjustments?.items;
  if (adjustmentItems?.minimum !== geometryProfiles.minimumValue || adjustmentItems?.maximum !== geometryProfiles.maximumValue)
    errors.push("PPJ adjustment bounds and canonical profile registry differ.");
  for (const [name, profile] of Object.entries(geometryProfiles.profiles ?? {})) {
    const guides = profile?.guides;
    const defaults = profile?.defaults;
    const parameters = profile?.parameters;
    if (typeof profile?.nativeToken !== "string" || profile.nativeToken.length === 0)
      errors.push(`Preset geometry profile ${name} has no native token.`);
    if (profile?.alias !== undefined && typeof profile.alias !== "boolean")
      errors.push(`Preset geometry profile ${name} has an invalid alias marker.`);
    if (!Array.isArray(guides) || !Array.isArray(defaults) || !Array.isArray(parameters) ||
        guides.length !== defaults.length || guides.length !== parameters.length)
      errors.push(`Preset geometry profile ${name} has inconsistent guide/default/parameter arrays.`);
    if (new Set(guides ?? []).size !== (guides ?? []).length)
      errors.push(`Preset geometry profile ${name} has duplicate guide names.`);
    for (const value of defaults ?? [])
      if (!Number.isInteger(value) || value < geometryProfiles.minimumValue || value > geometryProfiles.maximumValue)
        errors.push(`Preset geometry profile ${name} has an out-of-range default.`);
  }
  const preferredTokens = new Map();
  for (const [name, profile] of Object.entries(geometryProfiles.profiles ?? {})) {
    if (profile.alias) continue;
    if (preferredTokens.has(profile.nativeToken))
      errors.push(`Preset geometry profiles ${preferredTokens.get(profile.nativeToken)} and ${name} share native token ${profile.nativeToken}.`);
    else preferredTokens.set(profile.nativeToken, name);
  }
  for (const [name, profile] of Object.entries(geometryProfiles.profiles ?? {}))
    if (profile.alias && !preferredTokens.has(profile.nativeToken))
      errors.push(`Preset geometry alias ${name} has no preferred profile for native token ${profile.nativeToken}.`);
  const boundaryBehaviors = new Set(["fail-closed", "source-bound-only", "partial"]);
  for (const [index, boundary] of (value.authoredCompilerBoundaries ?? []).entries()) {
    if (!boundary || !boundaryBehaviors.has(boundary.behavior))
      errors.push(`Authored compiler boundary ${index} has an invalid behavior.`);
    for (const field of ["feature", "ppjPath", "sourceBound", "reason"])
      if (typeof boundary?.[field] !== "string" || boundary[field].length === 0)
        errors.push(`Authored compiler boundary ${index} is missing ${field}.`);
  }

  const editPlanSource = readFileSync(path.join(repo, "native/OfficeKit/src/OfficeKit.Codec/PptxEditPlanCodec.cs"), "utf8");
  const closedLeafSet = /leafKind is not \(([^\n]+)\)/u.exec(editPlanSource)?.[1] ?? "";
  const runtimeLeaves = [...closedLeafSet.matchAll(/"([^"]+)"/g)].map((match) => match[1]).sort();
  const registryLeaves = Object.keys(value.nativeLeafKinds ?? {}).sort();
  if (runtimeLeaves.length === 0) errors.push("Could not discover the closed native leaf set in PptxEditPlanCodec.cs.");
  if (JSON.stringify(runtimeLeaves) !== JSON.stringify(registryLeaves))
    errors.push("PptxEditPlanCodec native leaf kinds and capability registry entries differ.");
  const allowedLeafSurfaces = new Set(["native-leaf", "typed-operation"]);
  const allowedLeafValueTypes = new Set(["string", "number", "boolean", "asset"]);
  for (const [name, details] of Object.entries(value.nativeLeafKinds ?? {})) {
    if (!details || !allowedLeafSurfaces.has(details.surface)) errors.push(`Native leaf ${name} has an invalid surface.`);
    if (!allowedLeafValueTypes.has(details?.valueType)) errors.push(`Native leaf ${name} has an invalid valueType.`);
    for (const field of ["unit", "ppjLocation", "boundary"])
      if (typeof details?.[field] !== "string" || details[field].length === 0) errors.push(`Native leaf ${name} is missing ${field}.`);
  }

  const liveSource = readFileSync(path.join(repo, "src/live/adapters/powerpoint.mjs"), "utf8");
  const liveBlock = /POWERPOINT_LIVE_OPERATIONS\s*=\s*Object\.freeze\(\[([\s\S]*?)\]\)/u.exec(liveSource)?.[1] ?? "";
  const liveNames = [...liveBlock.matchAll(/"([a-z_]+)"/g)].map((match) => `powerpoint.${match[1]}`).sort();
  const declaredLive = [...(value.hostOnly ?? [])].sort();
  if (JSON.stringify(liveNames) !== JSON.stringify(declaredLive)) errors.push("PowerPoint Live operations and host-only registry entries differ.");
  for (const name of value.hostOnly ?? []) if (value.helpApis?.[name] != null) errors.push(`Host-only operation leaked into file authoring Help: ${name}`);
  validateSkillTree(errors);
  return errors;
}

function validateSkillTree(errors) {
  const root = path.join(repo, "skills/presentations/skills/presentations");
  const required = [
    "SKILL.md",
    "references/ppj.md",
    "references/fonts.md",
    "references/shapes.md",
    "references/text.md",
    "references/charts-and-tables.md",
    "references/media-and-layers.md",
    "references/motion.md",
    "references/components-and-templates.md",
    "references/imported-native-ref.md",
    "references/review-and-delivery.md",
    "references/scenarios/README.md",
    "references/scenarios/analysis-decision.md",
    "references/scenarios/business-proposal.md",
    "references/scenarios/management-report.md",
    "references/scenarios/academic-research.md",
    "references/scenarios/education-training.md",
    "references/scenarios/technical-engineering.md",
    "references/scenarios/brand-creative.md",
  ];
  const obsolete = [
    "tasks/create.md",
    "tasks/create-from-template.md",
    "tasks/edit-existing.md",
    "tasks/continue.md",
    "tasks/review-deliver.md",
    "references/authoring-plan.md",
    "references/primitives.md",
    "references/imported-capabilities.md",
    "references/source-continuation.md",
    "style_guidelines.md",
  ];
  for (const relative of required) if (!existsSync(path.join(root, relative))) errors.push(`Missing focused Presentation guidance: ${relative}`);
  for (const relative of obsolete) if (existsSync(path.join(root, relative))) errors.push(`Obsolete Presentation authority still exists: ${relative}`);
  const mainPath = path.join(root, "SKILL.md");
  if (!existsSync(mainPath)) return;
  const main = readFileSync(mainPath, "utf8");
  if (main.split(/\r?\n/u).length > 180) errors.push("Presentations SKILL.md exceeds the short-router budget of 180 lines.");
  for (const needle of ["references/ppj.md", "officekit ppj import", "references/review-and-delivery.md", "only public Presentation authoring language"]) {
    if (!main.includes(needle)) errors.push(`Presentations SKILL.md is missing required PPJ route text: ${needle}`);
  }
  for (const needle of ["tasks/create.md", "slide.compose", "presentation.editNativeLeaf", "references/primitives.md", "artifact_tool/api/"]) {
    if (main.includes(needle)) errors.push(`Legacy public Presentation route leaked into SKILL.md: ${needle}`);
  }
}

function renderManual(schema, registry, presetProfiles) {
  const schemaBytes = readFileSync(schemaPath);
  const registryBytes = readFileSync(registryPath);
  const minimumProgram = readFileSync(
    path.join(repo, "examples", "ppj", "minimum.ppj"),
    "utf8",
  ).trimEnd();
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
  const fieldOwnerRows = Object.entries(registry.ppjPathOwners).map(([field, details]) =>
    `| \`${field}\` | \`${details.owner}\` | \`${details.surface}\` | ${details.meaning} |`).join("\n");
  const ppjStateApiRows = Object.entries(registry.ppjStateApiPaths).map(([name, ppjPath]) =>
    `| \`${name}\` | \`${ppjPath}\` |`).join("\n");
  const nativeLeafRows = Object.entries(registry.nativeLeafKinds).map(([name, details]) =>
    `| \`${name}\` | \`${details.valueType}\` | \`${details.unit}\` | \`${details.ppjLocation}\` | \`${details.surface}\` | ${details.boundary} |`).join("\n");
  const authoredBoundaryRows = registry.authoredCompilerBoundaries.map((details) =>
    `| ${details.feature} | \`${details.ppjPath}\` | \`${details.behavior}\` | ${details.sourceBound} | ${details.reason} |`).join("\n");
  const presetGeometryRows = Object.entries(presetProfiles.profiles).map(([name, profile]) =>
    `| \`${name}\` | \`${profile.nativeToken}\`${profile.alias ? " (alias)" : ""} | ${profile.parameters.length ? profile.parameters.map(code).join(", ") : "none"} | ${profile.defaults.length ? `\`[${profile.defaults.join(", ")}]\`` : "native fixed geometry"} |`).join("\n");
  const definitionSections = Object.entries(schema.$defs).map(([name, definition]) => renderDefinition(name, definition, schema)).join("\n\n");
  const fieldCount = Object.values(schema.$defs).reduce((sum, definition) => sum + definitionProperties(definition, schema).properties.size, 0) + Object.keys(schema.properties).length;
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

## Capability map

The language contains ${fieldCount} documented root/definition fields and
${Object.keys(registry.nativeLeafKinds).length} closed source-edit leaf kinds.
The following owners prevent a runtime feature from existing without an Agent
route:

| PPJ path | Owner | Surface | Meaning |
| --- | --- | --- | --- |
${fieldOwnerRows}

## Typed page elements

| \`type\` | Required fields beyond \`type\` | Optional fields |
| --- | --- | --- |
${elementRows}

Simple text uses a string. Mixed formatting uses \`paragraphs[]\` and
\`runs[]\`; do not encode markup inside strings. Colors use theme references or
explicit typed color objects. Assets use relative URIs, exact MIME and SHA-256;
remote URLs and data-fetch instructions are invalid. Accessibility and rights
metadata travel with the asset or element.

### Element visibility and edit locks

Every typed element may declare \`hidden\` and \`locked\`. \`hidden: true\` keeps
the element and stable ID in PPJ but hides that object in the slide; it does not
hide the page or move the object in z-order. \`locked: true\` applies OfficeKit's
canonical edit lock for that object kind. It helps protect a finished visual
layer from accidental selection or movement, but it is not document security,
encryption, or an access-control boundary. Element array order remains the only
z-order.

Imported objects may change these fields only when their \`nativeRef\` issues
\`setHidden\` or \`setLocked\` for the current source revision. A partial or
extension-bearing native lock profile remains source-owned and does not become
an editable boolean.

## Preset geometry adjustments

\`shape.geometry.adjustments\` and preset \`image.mask.adjustments\` use one
complete ordered integer array. Omit it or use \`[]\` for the native preset
defaults. Percentage-like values use 100000 as 100%; angle values use 60000
units per degree. PPJ derives native guide names from the preset and never
exposes formula strings. The catalog contains ${Object.keys(presetProfiles.profiles).length}
PPJ names; availability does not make a shape useful, so choose geometry by
information purpose rather than novelty.

| Preset | Native token | Ordered parameter meaning | Native defaults |
| --- | --- | --- | --- |
${presetGeometryRows}

Imported shapes receive \`setGeometry\` only when their native adjustment list
is empty/default or contains the complete canonical guide order with literal
\`val N\` formulas. Formula-valued, partial, reordered, duplicated, or unknown
guides remain source-owned and reject geometry edits. Canonical imported
picture masks use the same rule and receive \`setImageMask\` only for
\`image.mask.adjustments\`; preset identity and picture topology remain fixed.

## Authored compiler availability

Schema validity proves that a program belongs to the PPJ language; it does not
claim that every declared state already has a source-free native writer. The
compiler owns ordinary text, shapes with preset geometry and solid fills,
images and native image backgrounds, charts, tables, connectors, groups,
placeholders, notes, comments, sections, custom shows, transitions, animations
and Morph subject to the explicit boundaries below. A boundary fails before
writing output; imported native content remains source-preserved.

| Feature | PPJ path | Authored behavior | Imported/source-bound behavior | Current reason |
| --- | --- | --- | --- | --- |
${authoredBoundaryRows}

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

### Imported presentation canvas

An imported \`design.canvas\` is editable only when its own \`nativeRef\` issues
\`setCanvas\` for \`canvas.width\` and \`canvas.height\`. Values stay in points.
Keep the nativeRef unchanged, edit one or both dimensions, then build and
re-import. The compiler changes only the native presentation canvas: it never
scales, reflows, crops or moves slide, layout, master or element coordinates.
Because every page is composed against that canvas, render and review every
page for exposed margins, clipping and changed balance after the edit.

### Imported raster/SVG fallback pairs

PowerPoint may store one visible picture as a raster compatibility fallback
plus an SVG used by modern hosts. PPJ projects these as two local assets:

- \`image.asset\` is the raster fallback;
- \`image.svgAsset\` is the true SVG member.

Change \`svgAsset\` only when that image's \`nativeRef\` issues \`replaceSvg\` for
\`image.svgAsset\`. Declare the replacement as a new content-addressed local
\`image/svg+xml\` asset, keep \`asset\`, the image ID and nativeRef unchanged,
then build and re-import. PPJ cannot add or remove the pair and does not invent
a raster fallback. Review the modern SVG render; when an older host matters,
record separately that its unchanged raster fallback may still show the prior
artwork.

### Closed imported leaf vocabulary

The compiler may issue the following bounded leaves after importing the exact
source PPTX. \`native-leaf\` values can appear in \`nativeRef.leaves[]\`;
\`typed-operation\` values are expressed by the typed PPJ element field named
below and still require the corresponding issued capability.

| Leaf kind | Value | Unit | PPJ location | Surface | Safety boundary |
| --- | --- | --- | --- | --- | --- |
${nativeLeafRows}

An imported native leaf has a stable opaque ID, a closed kind, the expected
source-value hash and its current scalar value. Change only \`value\`; never
invent an ID, kind or hash:

\`\`\`json
{
  "nativeRef": {
    "handle": "nr-…",
    "sourceSha256": "<64 lowercase hex>",
    "revision": "pptx-…",
    "objectHash": "<64 lowercase hex>",
    "capabilitySetSha256": "<64 lowercase hex>",
    "capabilities": [],
    "leaves": [
      {
        "id": "leaf-font-size-…",
        "kind": "fontSizePoints",
        "expectedHash": "<64 lowercase hex>",
        "value": 18
      }
    ]
  }
}
\`\`\`

For an imported deck, change \`pages[]\` order only when every moved page's
\`nativeRef\` advertises \`reorder\` with \`pageOrder\`. The projected page ID is
anchored to its source SlidePart, and unchanged page-local element IDs survive
the move. This is distinct from element \`reorder/zOrder\`, which changes the
stack inside one page. Do not combine page deletion and reorder in one build.
If modeled \`sections[]\` exist, update their \`pages\` arrays in the same program
so they still form one complete partition in the new presentation order.
Comments and custom shows keep referring to the same stable page IDs. An
opaque section graph receives no page-order capability and remains unchanged.

To reuse one complete imported page, find a page whose \`nativeRef\` advertises
\`duplicate\` with \`pageClone\`. Insert one fresh page immediately after that
source page with an empty \`elements\` array and
\`sourceClone: { page: "<source-page-id>", capability: "<issued-capability-id>" }\`.
Do not copy nativeRef, element state, layout, background, notes, transition,
animation or visibility into the pending clone. Build and re-import first; the
new SlidePart then projects as an ordinary source-bound page whose full typed
or opaque content can be inspected and edited through its newly issued
capabilities. One source page may have only one pending clone, and the same
build cannot also delete/reorder pages or change section/custom-show routes.

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
${minimumProgram}
\`\`\`

## Capability ownership

| Class | Current entries | Meaning |
| --- | ---: | --- |
${classRows}

The registry classifies legacy facade APIs while PPJ 2.0 converges. A
\`compiler-helper\` is not Agent syntax. PowerPoint Live operations remain in
the separate host-only list and never serialize into PPJ.

### Legacy state API mapping

Every facade API classified as persistent PPJ state names the language field
that replaces the method call. The maintainer rejects an unmapped entry:

| Legacy API | PPJ state |
| --- | --- |
${ppjStateApiRows}

## Complete schema field reference

This section is generated from every PPJ schema definition. It is exhaustive
for syntax; the focused references explain design judgment and workflows.

${definitionSections}

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
function renderDefinition(name, definition, schema) {
  const { properties, required } = definitionProperties(definition, schema);
  const summary = schemaSummary(definition);
  if (properties.size === 0)
    return `### \`${name}\`\n\n${summary}.`;
  const rows = [...properties.entries()].map(([field, fieldSchema]) =>
    `| \`${field}\` | ${required.has(field) ? "yes" : "no"} | ${schemaSummary(fieldSchema)} | ${schemaConstraints(fieldSchema)} |`).join("\n");
  return `### \`${name}\`\n\n${summary}.\n\n| Field | Required | Type or allowed values | Constraints |\n| --- | --- | --- | --- |\n${rows}`;
}
function definitionProperties(definition, schema) {
  const properties = new Map();
  const required = new Set();
  const visit = (fragment, visited = new Set()) => {
    if (!fragment || typeof fragment !== "object") return;
    if (fragment.$ref) {
      const name = fragment.$ref.split("/").at(-1);
      if (visited.has(name)) return;
      visit(schema.$defs[name], new Set([...visited, name]));
    }
    for (const part of fragment.allOf ?? []) visit(part, visited);
    for (const [field, value] of Object.entries(fragment.properties ?? {})) properties.set(field, value);
    for (const field of fragment.required ?? []) required.add(field);
  };
  visit(definition);
  return { properties, required };
}
function schemaSummary(value) {
  if (!value || typeof value !== "object") return "unknown";
  if (value.$ref) return `\`${value.$ref.split("/").at(-1)}\``;
  if (Object.hasOwn(value, "const")) return `literal ${code(JSON.stringify(value.const))}`;
  if (value.enum) return value.enum.map((item) => code(JSON.stringify(item))).join(" | ");
  if (value.oneOf) return value.oneOf.map(schemaSummary).join(" or ");
  if (value.allOf) return value.allOf.map(schemaSummary).join(" + ");
  if (value.type === "array") return `array of ${schemaSummary(value.items)}`;
  if (Array.isArray(value.type)) return value.type.map(code).join(" or ");
  if (value.type) return code(value.type);
  return "validated value";
}
function schemaConstraints(value) {
  if (!value || typeof value !== "object") return "none";
  const constraints = [];
  if (value.description) constraints.push(value.description);
  for (const [field, label] of [["minLength", "min chars"], ["maxLength", "max chars"], ["minItems", "min items"], ["maxItems", "max items"], ["minimum", "min"], ["exclusiveMinimum", ">"], ["maximum", "max"], ["exclusiveMaximum", "<"]])
    if (Object.hasOwn(value, field)) constraints.push(`${label} ${value[field]}`);
  if (value.pattern) constraints.push(`pattern ${code(value.pattern)}`);
  if (value.uniqueItems) constraints.push("unique items");
  return constraints.join("; ") || "none";
}
function digest(value) { return createHash("sha256").update(value).digest("hex"); }
function readJson(file) { return JSON.parse(readFileSync(file, "utf8")); }
function fail(errors) { for (const error of errors) process.stderr.write(`ERROR ${error}\n`); process.exit(1); }
