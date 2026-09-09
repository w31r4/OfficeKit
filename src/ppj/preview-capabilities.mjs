// Shared, filesystem-free projection of the registry's preview contract.
import { PREVIEW_STATUSES } from "./preview-diagnostics.mjs";

export function previewSchemaAt(schema, pointer) {
  if (typeof pointer !== "string" || !pointer.startsWith("#/")) throw new TypeError("Preview schema reference must be a local JSON pointer");
  let value = schema;
  for (const segment of pointer.slice(2).split("/").map((part) => part.replace(/~1/gu, "/").replace(/~0/gu, "~"))) {
    if (!value || typeof value !== "object" || !Object.hasOwn(value, segment)) throw new Error(`Missing preview schema reference: ${pointer}`);
    value = value[segment];
  }
  return value;
}

function ownProperties(schema, definition, seen = new Set()) {
  if (!definition || seen.has(definition)) return {};
  seen.add(definition);
  return Object.assign({}, definition.$ref ? ownProperties(schema, previewSchemaAt(schema, definition.$ref), seen) : {},
    ...(definition.allOf || []).map((part) => ownProperties(schema, part, seen)), definition.properties || {});
}

function requireSameNames(actual, expected, label) {
  if (JSON.stringify([...actual].sort()) !== JSON.stringify([...expected].sort())) throw new Error(`Preview ${label} must match the actual schema vocabulary`);
}

export function validatePreviewSupport(registry, schema) {
  const support = registry.previewSupport;
  if (support?.schema !== "office-kit/ppj-preview-support/v1") throw new Error("Missing or unknown preview support contract");
  if (typeof support.contract !== "string" || !support.contract) throw new Error("Missing preview support meaning");
  const definitions = schema.$defs.element.oneOf.map((item) => previewSchemaAt(schema, item.$ref));
  const types = definitions.map((definition) => ownProperties(schema, definition).type.const);
  const charts = ownProperties(schema, schema.$defs.chartElement).chartType.enum;
  requireSameNames(Object.keys(support.types), types, "element types");
  requireSameNames(Object.keys(support.charts), charts, "chart types");
  const owners = new Set(Object.values(registry.ppjPathOwners).map((entry) => entry.owner));
  for (const category of ["types", "charts", "variants", "fields", "sourceBound", "metadata", "factual"]) {
    const entries = support[category];
    if (!entries || Array.isArray(entries) || typeof entries !== "object") throw new Error(`Missing preview ${category}`);
    for (const [name, rule] of Object.entries(entries)) {
      const target = previewSchemaAt(schema, rule.schemaRef);
      if (category !== "metadata" && !PREVIEW_STATUSES.includes(rule.status)) throw new Error(`Invalid preview grade: ${category}.${name}`);
      if (category === "metadata" && Object.hasOwn(rule, "status")) throw new Error(`Metadata is not a visual support grade: ${name}`);
      if (category === "factual" && rule.severity !== "error") throw new Error(`Factual preview violations must be errors: ${name}`);
      if (!owners.has(rule.owner)) throw new Error(`Unknown preview owner: ${category}.${name}`);
      if (typeof rule.reason !== "string" || !rule.reason || typeof rule.test !== "string" || !rule.test.startsWith("test/")) throw new Error(`Preview rule needs a reason and regression: ${category}.${name}`);
      if (rule.stateHandler !== undefined && (typeof rule.stateHandler !== "string" || typeof rule.stateTest !== "string" || !rule.stateTest.startsWith("test/"))) throw new Error(`Preview state handler needs a regression: ${category}.${name}`);
      if (category === "types" && ownProperties(schema, target).type?.const !== name) throw new Error(`Wrong preview type owner: ${name}`);
      if (category === "charts" && !target.enum?.includes(name)) throw new Error(`Wrong preview chart owner: ${name}`);
      if (category === "variants" && (types.includes(name) || charts.includes(name))) throw new Error(`Preview variant is a schema type: ${name}`);
      // A promotion must identify the regression that proves full family state,
      // not merely continue to cite the declaration-presence test.
      if (rule.status === "supported" && (typeof rule.completeStateTest !== "string" || !rule.completeStateTest.startsWith("test/") || rule.completeStateTest === rule.test)) {
        throw new Error(`Supported preview rule needs complete-state evidence: ${category}.${name}`);
      }
    }
  }
  return support;
}

export function derivePreviewCapabilities(registry, schema) {
  const support = validatePreviewSupport(registry, schema);
  const summary = {
    schema: "office-kit/presentation-svg-preview-capabilities/v1",
    supported: [], partial: [], opaque: [], unavailable: [],
    variants: {}, sourceBound: {},
    source: "src/ppj/capability-registry.json#previewSupport",
    contract: support.contract,
  };
  for (const [name, rule] of Object.entries(support.types)) summary[rule.status].push(name);
  for (const [name, rule] of Object.entries(support.charts)) summary[rule.status].push(`chart:${name}`);
  for (const status of PREVIEW_STATUSES) summary[status].sort();
  for (const category of ["variants", "sourceBound"]) {
    for (const name of Object.keys(support[category]).sort()) summary[category][name] = support[category][name].status;
  }
  return summary;
}

export function assertPreviewCapabilitiesCurrent(actual, registry, schema) {
  const expected = derivePreviewCapabilities(registry, schema);
  if (JSON.stringify(actual) !== JSON.stringify(expected)) throw new Error("Preview capability summary is stale; run node scripts/generate-ppj-preview-capabilities.mjs");
  return expected;
}
