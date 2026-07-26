function sectionSnapshot(block, blockIndex, sectionOrdinal) {
  return {
    id: block.id,
    blockIndex,
    sectionOrdinal,
    name: block.name || "",
    editable: block.editable,
    breakType: block.breakType,
    orientation: block.orientation,
    pageSize: block.pageSize,
    margins: block.margins,
    columns: block.columns,
    pageNumbering: block.pageNumbering,
    lineNumbering: block.lineNumbering,
  };
}

export function boundedIndex(value, label) {
  const index = Number(value);
  if (!Number.isSafeInteger(index) || index < 0) throw new TypeError(`${label} must be a non-negative safe integer.`);
  return index;
}

export function sectionProjection(document) {
  let sectionOrdinal = 0;
  return document.blocks.flatMap((block, blockIndex) => {
    if (block.kind !== "section") return [];
    const snapshot = sectionSnapshot(block, blockIndex, sectionOrdinal);
    sectionOrdinal += 1;
    return [snapshot];
  });
}

export function selectCanonicalSection(document, sectionBlockIndex, label = "selected section") {
  const blockIndex = boundedIndex(sectionBlockIndex, "sectionBlockIndex");
  const block = document.blocks[blockIndex];
  if (!block) throw new Error("sectionBlockIndex is outside the imported document.");
  if (block.kind !== "section") throw new Error("sectionBlockIndex does not identify an imported section block.");
  if (!block.editable) throw new Error(`The ${label} is source-bound and read-only; its native w:sectPr graph is outside this workflow's canonical profile.`);
  if (document.resolve(block.id) !== block) throw new Error(`The ${label} locator did not resolve to the inspected object.`);
  const sectionOrdinal = document.blocks.slice(0, blockIndex + 1).filter((item) => item.kind === "section").length - 1;
  return { block, blockIndex, sectionOrdinal, snapshot: sectionSnapshot(block, blockIndex, sectionOrdinal) };
}

// This is deliberately not a general XML attribute parser. Source-bound
// section workflows only accept canonical WordprocessingML attributes on their
// own leaf, so another namespace (or a local-name collision) must fail before
// a masked residual could hide it.
export function wordAttributes(opening = "", label = "WordprocessingML leaf") {
  const tag = String(opening);
  const openingMatch = /^<w:[\w.-]+\b([\s\S]*?)\/>$/.exec(tag);
  if (!openingMatch) throw new Error(`${label} is not a canonical self-closing WordprocessingML leaf.`);
  const result = {};
  let remaining = openingMatch[1].trim();
  while (remaining) {
    const match = /^([:\w.-]+)="([^"]*)"\s*/.exec(remaining);
    if (!match) throw new Error(`${label} has unsupported XML attribute syntax.`);
    const [, qualifiedName, value] = match;
    const separator = qualifiedName.indexOf(":");
    if (separator <= 0 || qualifiedName.slice(0, separator) !== "w") {
      throw new Error(`${label} has a noncanonical attribute namespace: ${qualifiedName}.`);
    }
    const localName = qualifiedName.slice(separator + 1);
    if (!localName || Object.hasOwn(result, localName)) {
      throw new Error(`${label} has a duplicate or invalid w: attribute: ${qualifiedName}.`);
    }
    result[localName] = value;
    remaining = remaining.slice(match[0].length);
  }
  return result;
}

export function sectionProperties(xml) {
  return [...String(xml).matchAll(/<w:sectPr\b[\s\S]*?<\/w:sectPr>/g)].map((match, index) => ({
    index,
    xml: match[0],
    offset: match.index,
  }));
}
