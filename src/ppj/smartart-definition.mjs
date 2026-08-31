import { createHash } from "node:crypto";

const MIME = "application/vnd.officekit.smartart-definition+json";

export function copyOnWriteSmartArtDefinition(
  workspace,
  { elementId, definition, assetUriPrefix = "deck.assets/smartart" } = {},
) {
  if (!workspace?.root || typeof workspace.root !== "object") throw new Error("PPJ workspace must contain a parsed program.");
  if (typeof elementId !== "string" || !elementId) throw new Error("SmartArt copy-on-write requires an element ID.");
  if (!(definition instanceof Uint8Array) || definition.byteLength === 0 || definition.byteLength > 1024 * 1024) {
    throw new Error("SmartArt definition must contain 1 through 1048576 bytes.");
  }
  let parsed;
  try { parsed = JSON.parse(Buffer.from(definition).toString("utf8")); }
  catch { throw new Error("SmartArt definition must be strict UTF-8 JSON."); }
  if (parsed?.schema !== "office-kit/smartart-definition/v1" || !parsed.layout || !parsed.style || !parsed.colors) {
    throw new Error("SmartArt definition must use office-kit/smartart-definition/v1 with layout, style, and colors sections.");
  }
  if (!safeRelativeUri(`${assetUriPrefix}/definition.json`)) {
    throw new Error(`SmartArt asset URI prefix must stay relative to the PPJ directory: ${assetUriPrefix}`);
  }

  const root = structuredClone(workspace.root);
  const matches = [];
  for (const page of root.pages ?? []) collectElements(page.elements, elementId, matches);
  if (matches.length !== 1) throw new Error(`PPJ SmartArt element ${elementId} must resolve exactly once.`);
  const selected = matches[0];
  if (selected.type !== "smartArt" || typeof selected.definitionAsset !== "string") {
    throw new Error(`PPJ element ${elementId} is not a SmartArt instance backed by a definition asset.`);
  }
  const originalDefinitionAssetId = selected.definitionAsset;
  const original = (root.assets ?? []).find((asset) => asset?.id === originalDefinitionAssetId);
  if (!original || original.mimeType !== MIME) {
    throw new Error(`PPJ SmartArt definition asset ${originalDefinitionAssetId} is missing or has the wrong MIME type.`);
  }

  const digest = digestSha256(definition);
  const existing = (root.assets ?? []).find((asset) => asset?.mimeType === MIME && asset.sha256 === digest);
  const assetId = existing?.id ?? `smartart-definition-${digest}`;
  const uri = existing?.uri ?? `${assetUriPrefix.replace(/\/+$/u, "")}/${digest}.json`;
  if (!existing) {
    root.assets ??= [];
    root.assets.push({
      id: assetId,
      uri,
      mimeType: MIME,
      sha256: digest,
      rights: { status: "internal" },
      accessibility: { decorative: true },
    });
  }
  selected.definitionAsset = assetId;
  delete selected.layout;
  return Object.freeze({
    program: Buffer.from(`${JSON.stringify(root, null, 2)}\n`, "utf8"),
    root,
    originalDefinitionAssetId,
    definitionAssetId: assetId,
    asset: Object.freeze({
      id: assetId,
      uri,
      mimeType: MIME,
      sha256: digest,
      data: Buffer.from(definition),
      reused: Boolean(existing),
    }),
  });
}

function collectElements(elements, id, output) {
  for (const element of elements ?? []) {
    if (element?.id === id) output.push(element);
    collectElements(element?.elements, id, output);
  }
}

function safeRelativeUri(uri) {
  if (!uri || uri.includes("\\") || uri.includes("\0") || uri.startsWith("/")) return false;
  if (/^[A-Za-z][A-Za-z0-9+.-]*:/u.test(uri)) return false;
  return !uri.split("/").some((segment) => segment === "..");
}

function digestSha256(value) {
  return createHash("sha256").update(value).digest("hex");
}
