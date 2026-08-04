#!/usr/bin/env node

import { cp, mkdir, rm } from "node:fs/promises";
import path from "node:path";

const root = path.resolve(import.meta.dirname, "..");
const app = path.join(root, "apps", "powerpoint-addin");
const output = path.join(app, "dist");
const icon = path.join(root, "skills", "presentations", "skills", "presentations", "assets", "file-presentation.png");

await rm(output, { recursive: true, force: true });
await mkdir(path.join(output, "assets"), { recursive: true });
for (const filename of ["taskpane.html", "taskpane.css", "taskpane.js", "support.html"]) {
  await cp(path.join(app, "src", filename), path.join(output, filename));
}
await cp(icon, path.join(output, "assets", "officekit-powerpoint-32.png"));
await cp(icon, path.join(output, "assets", "officekit-powerpoint-80.png"));
