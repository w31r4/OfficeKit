import { readFile, writeFile } from "node:fs/promises";
import path from "node:path";
import { fileURLToPath } from "node:url";

import * as brands from "@fortawesome/free-brands-svg-icons";
import * as regular from "@fortawesome/free-regular-svg-icons";
import * as solid from "@fortawesome/free-solid-svg-icons";
import svgpath from "svgpath";

const root = path.resolve(path.dirname(fileURLToPath(import.meta.url)), "..");
const outputPath = path.join(root, "src", "ppj", "font-awesome-free-icons.json");
const checkOnly = process.argv.includes("--check");
const expectedVersion = "7.3.1";
const compareOrdinal = (left, right) => (left < right ? -1 : left > right ? 1 : 0);

const families = [
  ["fab", "@fortawesome/free-brands-svg-icons", brands],
  ["far", "@fortawesome/free-regular-svg-icons", regular],
  ["fas", "@fortawesome/free-solid-svg-icons", solid],
];

async function packageMetadata(packageName) {
  const packagePath = path.join(root, "node_modules", ...packageName.split("/"), "package.json");
  return JSON.parse(await readFile(packagePath, "utf8"));
}

function normalizePath(pathData, iconName) {
  const source = Array.isArray(pathData) ? pathData.join(" ") : pathData;
  const segments = [];
  svgpath(source)
    .abs()
    .unshort()
    .unarc()
    .round(4)
    .iterate((segment) => segments.push([...segment]));
  const normalized = segments
    .map((segment) => `${segment[0]}${segment.slice(1).join(" ")}`)
    .join("");
  const commands = new Set([...normalized.matchAll(/[A-Za-z]/g)].map((match) => match[0]));
  const unsupported = [...commands].filter((command) => !["M", "L", "C", "Z"].includes(command));
  if (unsupported.length > 0) {
    throw new Error(`${iconName}: unsupported normalized commands ${unsupported.join(", ")}`);
  }
  if (normalized.length === 0 || normalized.length > 128_000) {
    throw new Error(`${iconName}: normalized path length ${normalized.length} is outside bounds`);
  }
  return normalized;
}

const icons = {};
const packages = [];

for (const [prefix, packageName, module] of families) {
  const metadata = await packageMetadata(packageName);
  if (metadata.version !== expectedVersion) {
    throw new Error(`${packageName}: expected ${expectedVersion}, found ${metadata.version}`);
  }
  packages.push({ name: packageName, version: metadata.version, license: metadata.license });

  const canonical = new Map();
  for (const value of Object.values(module)) {
    if (
      value &&
      typeof value === "object" &&
      value.prefix === prefix &&
      typeof value.iconName === "string" &&
      Array.isArray(value.icon)
    ) {
      canonical.set(value.iconName, value);
    }
  }

  for (const [name, value] of [...canonical.entries()].sort(([a], [b]) => compareOrdinal(a, b))) {
    const [width, height, , , pathData] = value.icon;
    const key = `${prefix}:${name}`;
    icons[key] = {
      width,
      height,
      path: normalizePath(pathData, key),
    };
  }
}

const sortedIcons = Object.fromEntries(Object.entries(icons).sort(([a], [b]) => compareOrdinal(a, b)));
const catalog = {
  schema: "office-kit/ppj-icon-catalog/v1",
  source: {
    name: "Font Awesome Free",
    version: expectedVersion,
    homepage: "https://fontawesome.com/",
    license: "CC-BY-4.0",
    codeLicense: "MIT",
    packages,
  },
  icons: sortedIcons,
};

const serialized = `${JSON.stringify(catalog)}\n`;

if (checkOnly) {
  const current = await readFile(outputPath, "utf8");
  if (current !== serialized) {
    throw new Error(`PPJ icon catalog is stale: run node ${path.relative(root, fileURLToPath(import.meta.url))}`);
  }
  process.stdout.write(`PPJ icon catalog current: ${Object.keys(sortedIcons).length} icons\n`);
} else {
  await writeFile(outputPath, serialized);
  process.stdout.write(`Generated ${path.relative(root, outputPath)} with ${Object.keys(sortedIcons).length} icons\n`);
}
