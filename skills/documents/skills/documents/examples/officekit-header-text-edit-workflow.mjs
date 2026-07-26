import path from "node:path";
import { pathToFileURL } from "node:url";

import {
  editBoundPageFurnitureText,
  pageFurnitureCliOutput,
  parsePageFurnitureTextEditCli,
} from "./officekit-page-furniture-text-edit.mjs";

export async function editImportedHeaderText(options) {
  return editBoundPageFurnitureText({ ...options, kind: "header" });
}

const entry = process.argv[1] ? pathToFileURL(path.resolve(process.argv[1])).href : "";
if (entry === import.meta.url) {
  const result = await editImportedHeaderText(parsePageFurnitureTextEditCli(process.argv.slice(2)));
  console.log(JSON.stringify(pageFurnitureCliOutput(result)));
}
